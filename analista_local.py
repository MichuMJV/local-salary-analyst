import os
import re
import json
import hashlib
import threading
import unicodedata
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import io
import contextlib

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import ollama

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
APP_TITLE = "Gemma - Analista Financiero (Code Interpreter)"
APP_GEOMETRY = "1380x820"
MEMORY_FILE = "financial_assistant_memory.json"

DEFAULT_MODEL = "gemma4:e4b" # Asegúrate de usar el modelo que tienes descargado

MAX_CHAT_HISTORY = 8
MAX_SAMPLE_ROWS = 8
MAX_TEXT_CONTEXT_CHARS = 12000

# =========================================================
# UTILIDADES
# =========================================================
def normalize_text(value):
    if value is None: return ""
    value = str(value).strip().lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return value

def shorten_text(text, max_chars=MAX_TEXT_CONTEXT_CHARS):
    text = str(text)
    if len(text) <= max_chars: return text
    return text[:max_chars] + "\n...[contexto recortado por tamaño]..."

def make_signature(filepath, sheet_name, columns):
    raw = f"{filepath}|{sheet_name}|{'|'.join(map(str, columns))}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()

def get_chunk_message_dict(chunk):
    try:
        if isinstance(chunk, dict): return chunk.get("message", {}) or {}
    except: pass
    try:
        msg = getattr(chunk, "message", None)
        if msg is None: return {}
        if isinstance(msg, dict): return msg
        return {
            "role": getattr(msg, "role", None),
            "content": getattr(msg, "content", "") or "",
            "thinking": getattr(msg, "thinking", "") or "",
        }
    except: return {}

# =========================================================
# MEMORIA
# =========================================================
class MemoryManager:
    def __init__(self, filepath=MEMORY_FILE):
        self.filepath = filepath
        self.data = self._load()

    def _load(self):
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"Advertencia al cargar memoria: {e}")
        return {"datasets": {}, "global": {"last_model": DEFAULT_MODEL, "recent_history": []}}

    def save(self):
        try:
            with open(self.filepath, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except PermissionError:
            # Si OneDrive u otro programa bloquea el archivo, lo omitimos en vez de crashear la app.
            print("Advertencia: El archivo de memoria está bloqueado (posiblemente por OneDrive). Se omitió el guardado en este ciclo.")
        except Exception as e:
            print(f"Error inesperado al guardar memoria: {e}")

    def get_last_model(self): return self.data.get("global", {}).get("last_model", DEFAULT_MODEL)
    
    def set_last_model(self, model_name):
        self.data.setdefault("global", {})["last_model"] = model_name
        self.save()

    def get_recent_history(self): return self.data.get("global", {}).get("recent_history", [])
    
    def set_recent_history(self, history):
        self.data.setdefault("global", {})["recent_history"] = history[-20:]
        self.save()

    def get_dataset(self, signature): return self.data.get("datasets", {}).get(signature, {})
    
    def update_dataset(self, signature, payload):
        self.data.setdefault("datasets", {})[signature] = payload
        self.save()

# =========================================================
# PROCESADOR DE EXCEL (AHORA MUCHO MÁS LIGERO)
# =========================================================
class ExcelProcessor:
    def __init__(self, filepath, sheet_name):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.sheet_name = sheet_name
        self.df = None
        self.signature = None
        self.precomputed = {}

    def load(self):
        ext = os.path.splitext(self.filepath)[1].lower()
        engine = "openpyxl" if ext == ".xlsx" else "xlrd" if ext == ".xls" else None

        self.df = pd.read_excel(self.filepath, sheet_name=self.sheet_name, engine=engine).fillna("")
        self.df.columns = [str(c).strip() for c in self.df.columns]
        self.signature = make_signature(self.filepath, self.sheet_name, self.df.columns.tolist())
        self._compute_predefined_metrics()

    def _compute_predefined_metrics(self):
        df = self.df.copy()
        self.precomputed = {
            "rows": len(df),
            "columns_count": len(df.columns),
            "sample_rows": df.head(MAX_SAMPLE_ROWS).to_string(index=False),
            "numeric_columns": [col for col in df.columns if pd.api.types.is_numeric_dtype(pd.to_numeric(df[col], errors='ignore'))]
        }

    def build_human_summary_text(self):
        lines = [
            f"📁 Archivo: {self.filename} | Hoja: {self.sheet_name}",
            f"📐 Dimensiones: {self.precomputed['rows']} filas x {self.precomputed['columns_count']} columnas",
            f"📌 Columnas: {', '.join(self.df.columns)}",
            f"🔢 Numéricas: {', '.join(self.precomputed['numeric_columns'])}",
            "\n🔍 Muestra de datos:\n" + self.precomputed['sample_rows']
        ]
        return "\n".join(lines)

    def build_llm_context(self):
        # Contexto estructurado para que Gemma escriba el código basado en esto
        return (
            f"Archivo: {self.filename}\n"
            f"Columnas disponibles: {list(self.df.columns)}\n"
            f"Columnas numéricas detectadas: {self.precomputed['numeric_columns']}\n"
            f"Muestra de las primeras {MAX_SAMPLE_ROWS} filas:\n{self.precomputed['sample_rows']}\n"
        )

# =========================================================
# APP PRINCIPAL (INTERFAZ Y EJECUCIÓN AGÉNTICA)
# =========================================================
class FinancialAssistantApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(APP_GEOMETRY)
        self.root.configure(bg="#161e2a")

        self.memory = MemoryManager()
        self.processor = None
        self.df = None
        self.history = self.memory.get_recent_history() or []
        
        self.live_answer_started = False
        self.last_chunk = ""

        self.setup_ui()
        self.model_var.set(self.memory.get_last_model() or DEFAULT_MODEL)
        
        # PROMPT DE SISTEMA ORIENTADO AL CODE INTERPRETER
        self.system_prompt_text = (
            "Eres un Científico de Datos y Analista Financiero Avanzado. "
            "Tienes acceso directo a un DataFrame de pandas cargado en la variable `df` y a la librería `plt` (matplotlib.pyplot).\n\n"
            "REGLAS CRÍTICAS PARA RESPONDER:\n"
            "1. Si el usuario pide un cálculo, variación, cruce de datos o gráfico, NO intentes adivinar la respuesta. "
            "DEBES generar un bloque de código en Python que resuelva la petición.\n"
            "2. El código debe estar estrictamente dentro de ```python y ```.\n"
            "3. Para devolver resultados en texto/números usa `print()`. El sistema ejecutará el código y te devolverá lo que imprimas.\n"
            "4. Para generar gráficos usa `plt.plot()`, `plt.bar()`, etc., y termina siempre con `plt.show()`.\n"
            "5. Si el usuario hace una pregunta general que no requiere matemáticas (ej. 'hola', 'qué puedes hacer'), responde con texto normal sin código."
        )

        self.append_to_chat("Sistema", "📊 Asistente iniciado con capacidades de Code Interpreter. Carga un Excel para comenzar.")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        # (El mismo diseño UI limpio que ya tenías)
        top_frame = tk.Frame(self.root, bg="#2c3e50", pady=8)
        top_frame.pack(fill=tk.X)

        tk.Label(top_frame, text="AInalista Financiero - Code Interpreter", fg="white", bg="#2c3e50", font=("Arial", 14, "bold")).pack(side=tk.LEFT, padx=16)
        
        tk.Label(top_frame, text="Modelo:", fg="white", bg="#2c3e50", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(12, 6))
        self.model_var = tk.StringVar(value=DEFAULT_MODEL)
        tk.Entry(top_frame, textvariable=self.model_var, width=18, font=("Consolas", 10)).pack(side=tk.LEFT, padx=(0, 10))

        self.show_thinking_var = tk.BooleanVar(value=False)
        tk.Checkbutton(top_frame, text="Mostrar thinking", variable=self.show_thinking_var, bg="#2c3e50", fg="white", selectcolor="#2c3e50", activebackground="#2c3e50", activeforeground="white", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(0, 12))
        
        tk.Button(top_frame, text="Limpiar thinking", command=self.clear_thinking_panel, bg="#7f8c8d", fg="white", font=("Arial", 10, "bold"), relief=tk.FLAT).pack(side=tk.LEFT, padx=(0, 12))
        tk.Button(top_frame, text="Cargar Excel", command=self.load_excel, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief=tk.FLAT).pack(side=tk.RIGHT, padx=20)

        main_pane = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, bg="#161e2a")
        main_pane.pack(expand=True, fill=tk.BOTH, padx=12, pady=12)

        left_frame = tk.Frame(main_pane, bg="#161e2a")
        main_pane.add(left_frame, minsize=760)

        self.chat_display = scrolledtext.ScrolledText(left_frame, wrap=tk.WORD, bg="white", font=("Consolas", 11), state=tk.DISABLED)
        self.chat_display.pack(expand=True, fill=tk.BOTH, padx=(8, 8), pady=(8, 8))

        bottom_frame = tk.Frame(left_frame, bg="#2980b9")
        bottom_frame.pack(fill=tk.X, padx=8, pady=(0, 8))

        self.input_box = tk.Text(bottom_frame, height=4, font=("Consolas", 11))
        self.input_box.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(0, 10))
        self.input_box.bind("<Return>", self.handle_return)

        self.btn_send = tk.Button(bottom_frame, text="Enviar", command=self.send_message, bg="#2980b9", fg="white", font=("Arial", 11, "bold"), width=12, relief=tk.FLAT)
        self.btn_send.pack(side=tk.RIGHT, fill=tk.Y)

        right_frame = tk.Frame(main_pane, bg="#ecf0f1", width=430)
        main_pane.add(right_frame, minsize=320)

        thinking_header = tk.Frame(right_frame, bg="#34495e")
        thinking_header.pack(fill=tk.X)
        tk.Label(thinking_header, text="Thinking Process (debug)", fg="white", bg="#34495e", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=10, pady=8)
        self.lbl_thinking_status = tk.Label(thinking_header, text="Inactivo", fg="#dfe6e9", bg="#34495e", font=("Arial", 10))
        self.lbl_thinking_status.pack(side=tk.RIGHT, padx=10)

        self.thinking_display = scrolledtext.ScrolledText(right_frame, wrap=tk.WORD, bg="#fdfdfd", font=("Consolas", 10), state=tk.DISABLED)
        self.thinking_display.pack(expand=True, fill=tk.BOTH, padx=8, pady=8)

    # (Funciones UI base sin cambios drásticos)
    def append_to_chat(self, sender, text):
        self.chat_display.config(state=tk.NORMAL)
        colors = {"Tú": "#2980b9", "Gemma": "#27ae60", "Sistema": "#8e44ad", "Consola Python": "#d35400", "Error": "#c0392b"}
        color = colors.get(sender, "#000000")
        self.chat_display.tag_config(sender, foreground=color, font=("Consolas", 11, "bold"))
        self.chat_display.insert(tk.END, f"{sender}:\n", sender)
        self.chat_display.insert(tk.END, f"{text}\n\n")
        self.chat_display.yview(tk.END)
        self.chat_display.config(state=tk.DISABLED)

    def clear_thinking_panel(self):
        self.thinking_display.config(state=tk.NORMAL)
        self.thinking_display.delete("1.0", tk.END)
        self.thinking_display.config(state=tk.DISABLED)
        self.lbl_thinking_status.config(text="Inactivo")

    def set_thinking_status(self, text): self.lbl_thinking_status.config(text=text)

    def handle_return(self, event):
        if not event.state & 0x0001:
            self.send_message()
            return "break"

    def _start_live_answer(self, sender="Gemma", color="#27ae60"):
        if self.live_answer_started: return
        self.live_answer_started = True
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.tag_config(sender, foreground=color, font=("Consolas", 11, "bold"))
        self.chat_display.insert(tk.END, f"{sender}:\n", sender)
        self.chat_display.config(state=tk.DISABLED)
        self.chat_display.yview(tk.END)

    def _append_live_answer(self, chunk_text):
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.insert(tk.END, chunk_text)
        self.chat_display.yview(tk.END)
        self.chat_display.config(state=tk.DISABLED)

    def _finish_live_answer(self):
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.insert(tk.END, "\n\n")
        self.chat_display.config(state=tk.DISABLED)
        self.chat_display.yview(tk.END)
        self.live_answer_started = False
        self.last_chunk = ""

    def _update_live_thinking(self, text):
        self.thinking_display.config(state=tk.NORMAL)
        self.thinking_display.delete("1.0", tk.END)
        self.thinking_display.insert(tk.END, text)
        self.thinking_display.config(state=tk.DISABLED)
        self.thinking_display.yview(tk.END)

    # Carga de Excel simplificada
    def load_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath: return
        self.btn_send.config(state=tk.DISABLED)
        self.append_to_chat("Sistema", f"📥 Cargando archivo: {os.path.basename(filepath)} ...")
        threading.Thread(target=self._load_excel_worker, args=(filepath,), daemon=True).start()

    def _load_excel_worker(self, filepath):
        try:
            excel_file = pd.ExcelFile(filepath)
            sheet = excel_file.sheet_names[0] # Tomamos la primera hoja por defecto para agilizar
            
            self.processor = ExcelProcessor(filepath, sheet)
            self.processor.load()
            self.df = self.processor.df
            
            summary = self.processor.build_human_summary_text()
            self.root.after(0, lambda: self.append_to_chat("Sistema", "🧠 Contexto del Excel cargado.\n\n" + summary))
        except Exception as e:
            self.root.after(0, lambda: self.append_to_chat("Error", f"Fallo al cargar: {str(e)}"))
        finally:
            self.root.after(0, lambda: self.btn_send.config(state=tk.NORMAL, text="Enviar"))

    def send_message(self):
        user_text = self.input_box.get("1.0", tk.END).strip()
        if not user_text: return
        
        self.input_box.delete("1.0", tk.END)
        self.append_to_chat("Tú", user_text)
        self.btn_send.config(state=tk.DISABLED, text="Pensando...")

        # Añadir al historial
        self.history.append({"role": "user", "content": user_text})
        self.history = self.history[-MAX_CHAT_HISTORY:]

        threading.Thread(target=self._orchestrate_ai_workflow, daemon=True).start()

    # =========================================================
    # EL MOTOR AGÉNTICO (Local Code Interpreter)
    # =========================================================
    def _orchestrate_ai_workflow(self):
        try:
            # 1. Llamada inicial a Gemma
            ai_reply = self._stream_ollama_call()
            self.history.append({"role": "assistant", "content": ai_reply})

            # 2. Buscar si Gemma decidió escribir código Python
            code_blocks = re.findall(r'```python(.*?)```', ai_reply, re.DOTALL)
            
            if code_blocks:
                for idx, code in enumerate(code_blocks):
                    self.root.after(0, lambda: self.append_to_chat("Sistema", f"⚙️ Ejecutando código generado por IA..."))
                    
                    # 3. Ejecutar el código en el entorno local (Sandbox)
                    success, output, fig = self._run_python_sandbox(code.strip())
                    
                    if fig:
                        self.root.after(0, lambda f=fig: self._show_figure(f, "Gráfico generado dinámicamente."))
                    
                    if output.strip():
                        self.root.after(0, lambda o=output: self.append_to_chat("Consola Python", o))
                    
                    # 4. Interactive RAG: Devolver el resultado a Gemma para que lo explique
                    if success:
                        feedback_prompt = f"El código se ejecutó exitosamente. Salida en consola:\n{output}\nPor favor, resume o explica este resultado al usuario en lenguaje natural."
                    else:
                        feedback_prompt = f"El código falló con este error:\n{output}\nInforma al usuario del error y sugiere cómo reformular la pregunta."
                    
                    self.history.append({"role": "user", "content": feedback_prompt})
                    
                    # Segunda llamada a Gemma para la conclusión
                    final_reply = self._stream_ollama_call(is_feedback=True)
                    self.history.append({"role": "assistant", "content": final_reply})

        except Exception as e:
            self.root.after(0, lambda: self.append_to_chat("Error", f"Error en el flujo: {str(e)}"))
        finally:
            self.memory.set_recent_history(self.history)
            self.root.after(0, lambda: self.btn_send.config(state=tk.NORMAL, text="Enviar"))

    def _run_python_sandbox(self, code_string):
        """Ejecuta el código dinámico generado por la IA en un entorno seguro y captura la salida."""
        if self.processor is None:
            return False, "Error: No hay DataFrame cargado.", None

        output_buffer = io.StringIO()
        
        # Preparamos matplotlib para que no bloquee los hilos de Tkinter
        plt.clf() 
        custom_show_called = {"called": False}
        
        # Mockeamos plt.show() para interceptar el gráfico en vez de que abra una ventana insegura
        def mock_show(*args, **kwargs):
            custom_show_called["called"] = True
            
        # El entorno donde correrá el código (Tiene el df real y las librerías)
        env = {
            'df': self.processor.df, 
            'pd': pd, 
            'plt': plt
        }
        env['plt'].show = mock_show

        try:
            # Ejecutamos el código y capturamos todos los print() que haya hecho la IA
            with contextlib.redirect_stdout(output_buffer):
                exec(code_string, env)
            
            # Verificamos si se generó algún gráfico
            fig = None
            if custom_show_called["called"] or plt.gcf().get_axes():
                fig = plt.gcf()
                
            return True, output_buffer.getvalue(), fig
        except Exception as e:
            return False, str(e), None

    def _stream_ollama_call(self, is_feedback=False):
        model_name = self.model_var.get().strip() or DEFAULT_MODEL
        
        messages = [{"role": "system", "content": self.system_prompt_text}]
        if self.processor:
            messages.append({"role": "system", "content": "Metadatos del DataFrame actual:\n" + self.processor.build_llm_context()})

        messages.extend(self.history)

        show_thinking = bool(self.show_thinking_var.get())
        thinking_text = ""
        answer_text = ""

        self.root.after(0, self.clear_thinking_panel)
        self.root.after(0, lambda: self.set_thinking_status("Analizando..." if not is_feedback else "Concluyendo..."))
        
        # Si es retroalimentación (is_feedback), lo mostramos como "Sistema" o "Gemma (Conclusión)"
        sender_name = "Gemma" if not is_feedback else "Gemma (Análisis de resultados)"
        self.root.after(0, lambda: self._start_live_answer(sender=sender_name))

        try:
            stream = ollama.chat(
                model=model_name,
                messages=messages,
                think=show_thinking,
                stream=True,
                options={"temperature": 0.1, "top_p": 0.9} # Temperatura baja para que el código sea preciso
            )

            update_counter = 0
            for chunk in stream:
                msg = get_chunk_message_dict(chunk)
                
                # Procesar Thinking
                chunk_thinking = msg.get("thinking", "") or ""
                if chunk_thinking:
                    thinking_text += chunk_thinking
                    if len(thinking_text) > 4000: thinking_text = thinking_text[-4000:]
                    if show_thinking and update_counter % 8 == 0:
                        self.root.after(0, lambda t=thinking_text: self._update_live_thinking(t))
                        self.root.after(0, lambda: self.set_thinking_status("Thinking..."))

                # Procesar Contenido (Respuesta)
                chunk_content = msg.get("content", "") or ""
                if chunk_content and chunk_content != self.last_chunk:
                    self.last_chunk = chunk_content
                    answer_text += chunk_content
                    self.root.after(0, lambda t=chunk_content: self._append_live_answer(t))
                    
                update_counter += 1

        except Exception as e:
            answer_text = f"[Error de conexión con Ollama: {e}]"
            self.root.after(0, lambda: self._append_live_answer(answer_text))

        self.root.after(0, lambda: self.set_thinking_status("Completado"))
        self.root.after(0, self._finish_live_answer)
        
        return answer_text

    def _show_figure(self, fig, msg):
        # Esta función corre en el hilo principal y levanta el gráfico de forma segura
        win = tk.Toplevel(self.root)
        win.title("Gráfico generado por IA")
        win.geometry("800x500")

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill=tk.BOTH)

    def on_close(self):
        try:
            self.memory.set_recent_history(self.history)
            self.memory.set_last_model(self.model_var.get().strip() or DEFAULT_MODEL)
            self.memory.save()
        except: pass
        self.root.destroy()

# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = FinancialAssistantApp(root)
    root.mainloop()