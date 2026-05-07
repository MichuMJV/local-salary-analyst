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

# --- FIX DE MATPLOTLIB ---
import matplotlib
matplotlib.use('Agg') # Obliga a matplotlib a renderizar en memoria y no chocar con los hilos de Tkinter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import ollama

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
APP_TITLE = "AInalista Financiero - Code Interpreter"
APP_GEOMETRY = "1380x820"
MEMORY_FILE = "financial_assistant_memory.json"

DEFAULT_MODEL = "gemma4:e4b"

MAX_CHAT_HISTORY = 10
MAX_SAMPLE_ROWS = 10

# =========================================================
# UTILIDADES
# =========================================================
def normalize_text(value):
    if value is None: return ""
    value = str(value).strip().lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return value

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
                print(f"Aviso al cargar memoria: {e}")
        return {"datasets": {}, "global": {"last_model": DEFAULT_MODEL, "recent_history": []}}

    def save(self):
        try:
            with open(self.filepath, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except PermissionError:
            print("Aviso: Archivo de memoria bloqueado (ej. por OneDrive). Se omitió el guardado.")
        except Exception as e:
            print(f"Error al guardar memoria: {e}")

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
# PROCESADOR DE EXCEL Y AUTODETECCIÓN
# =========================================================
class ExcelProcessor:
    def __init__(self, filepath, sheet_name, column_definitions=None):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.sheet_name = sheet_name
        self.df = None
        self.signature = None
        
        self.column_definitions = column_definitions or {}
        self.column_roles = {}
        self.column_meanings = {}
        self.precomputed = {}

    def load(self):
        ext = os.path.splitext(self.filepath)[1].lower()
        engine = "openpyxl" if ext == ".xlsx" else "xlrd" if ext == ".xls" else None
        self.df = pd.read_excel(self.filepath, sheet_name=self.sheet_name, engine=engine).fillna("")
        self.df.columns = [str(c).strip() for c in self.df.columns]
        self.signature = make_signature(self.filepath, self.sheet_name, self.df.columns.tolist())
        
        # 1. Cargar definiciones de memoria si existen
        for col in self.df.columns:
            if col in self.column_definitions:
                self.column_roles[col] = self.column_definitions[col].get("role", "")
                self.column_meanings[col] = self.column_definitions[col].get("meaning", "")
            else:
                self.column_roles[col] = ""
                self.column_meanings[col] = ""
                
        # 2. Autodetectar roles
        self._infer_roles()
        self._compute_predefined_metrics()

    def _infer_roles(self):
        for col in self.df.columns:
            if self.column_roles.get(col):
                continue
                
            series = self.df[col]
            if pd.api.types.is_datetime64_any_dtype(series):
                self.column_roles[col] = "fecha"
                continue
                
            num_series = pd.to_numeric(series, errors='coerce')
            if num_series.notna().sum() > len(series) * 0.4:
                self.column_roles[col] = "valores"
                continue
                
            unique_count = series.nunique()
            if unique_count <= 25 or unique_count < len(series) * 0.15:
                self.column_roles[col] = "categoría principal"
            else:
                self.column_roles[col] = "texto descriptivo"

    def apply_user_definitions(self, definitions):
        if not definitions: return
        for col, meta in definitions.items():
            if col in self.df.columns:
                self.column_roles[col] = meta.get("role", "unknown")
                self.column_meanings[col] = meta.get("meaning", "")
        self._compute_predefined_metrics()

    def _compute_predefined_metrics(self):
        self.precomputed = {
            "rows": len(self.df),
            "columns_count": len(self.df.columns),
            "sample_rows": self.df.head(MAX_SAMPLE_ROWS).to_string(index=False),
        }

    def build_llm_context(self):
        lines = [f"Archivo: {self.filename} (Hoja: {self.sheet_name})", "DICCIONARIO DE DATOS (Metadata):"]
        for col in self.df.columns:
            role = self.column_roles.get(col, "unknown")
            meaning = self.column_meanings.get(col, "")
            lines.append(f"- {col}: Tipo={role} | Significado={meaning}")
        
        lines.append(f"\nMuestra de datos ({MAX_SAMPLE_ROWS} filas):\n{self.precomputed['sample_rows']}")
        return "\n".join(lines)

    def build_memory_payload(self):
        return {
            "file": self.filename,
            "sheet": self.sheet_name,
            "signature": self.signature,
            "column_definitions": {
                col: {"role": self.column_roles.get(col, ""), "meaning": self.column_meanings.get(col, "")}
                for col in self.df.columns
            },
            "updated_at": datetime.now().isoformat()
        }

# =========================================================
# DIÁLOGOS DE INTERFAZ
# =========================================================
class SheetSelectorDialog(tk.Toplevel):
    def __init__(self, parent, sheets):
        super().__init__(parent)
        self.title("Seleccionar hoja")
        self.geometry("380x320")
        self.resizable(False, False)
        self.selected_sheet = None
        self.transient(parent)
        self.grab_set()

        lbl = tk.Label(self, text="El archivo tiene varias hojas.\nSelecciona la hoja a analizar:", font=("Arial", 11, "bold"))
        lbl.pack(pady=10)

        self.listbox = tk.Listbox(self, font=("Consolas", 11))
        self.listbox.pack(expand=True, fill=tk.BOTH, padx=15, pady=10)
        for s in sheets: self.listbox.insert(tk.END, s)
        if sheets: self.listbox.selection_set(0)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Aceptar", command=self.on_ok, bg="#2980b9", fg="white", width=12).pack(side=tk.LEFT, padx=8)
        tk.Button(btn_frame, text="Cancelar", command=self.on_cancel, bg="#7f8c8d", fg="white", width=12).pack(side=tk.LEFT, padx=8)
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def on_ok(self):
        selection = self.listbox.curselection()
        if selection: self.selected_sheet = self.listbox.get(selection[0])
        self.destroy()

    def on_cancel(self):
        self.selected_sheet = None
        self.destroy()

class ColumnDefinitionDialog(tk.Toplevel):
    def __init__(self, parent, processor):
        super().__init__(parent)
        self.title("Diccionario de Datos")
        self.geometry("980x600")
        self.configure(bg="white")
        self.processor = processor
        self.result = {}
        self.transient(parent)
        self.grab_set()

        header = tk.Label(self, text="El sistema ha detectado automáticamente el tipo de dato.\nEscribe el diccionario de datos (Significado) de las columnas para la IA.", bg="white", fg="#1f2937", font=("Arial", 11, "bold"), justify=tk.LEFT)
        header.pack(anchor="w", padx=12, pady=10)

        container = tk.Frame(self, bg="white")
        container.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        canvas = tk.Canvas(container, bg="white", highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg="white")

        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        hdr = tk.Frame(self.scrollable_frame, bg="#ecf0f1")
        hdr.pack(fill=tk.X, pady=(0, 5))
        tk.Label(hdr, text="Columna", width=28, anchor="w", bg="#ecf0f1", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", padx=4, pady=4)
        tk.Label(hdr, text="Tipo Detectado (Automático)", width=28, anchor="w", bg="#ecf0f1", font=("Arial", 10, "bold")).grid(row=0, column=1, sticky="w", padx=4, pady=4)
        tk.Label(hdr, text="Diccionario de Datos (Significado)", width=42, anchor="w", bg="#ecf0f1", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky="w", padx=4, pady=4)

        self.widgets = {}
        for col in self.processor.df.columns:
            row = tk.Frame(self.scrollable_frame, bg="white")
            row.pack(fill=tk.X, pady=2)

            tk.Label(row, text=col, width=28, anchor="w", bg="white", font=("Consolas", 10)).grid(row=0, column=0, sticky="w", padx=4)

            inferred_role = self.processor.column_roles.get(col, "desconocido")
            tk.Label(row, text=inferred_role.upper(), width=28, anchor="w", bg="white", fg="#2980b9", font=("Consolas", 9, "bold")).grid(row=0, column=1, sticky="w", padx=4)

            meaning_var = tk.StringVar(value=self.processor.column_meanings.get(col, ""))
            ent = tk.Entry(row, textvariable=meaning_var, width=58, font=("Consolas", 10))
            ent.grid(row=0, column=2, sticky="we", padx=4)

            self.widgets[col] = {"role": inferred_role, "meaning_var": meaning_var}

        btn_frame = tk.Frame(self, bg="white")
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Button(btn_frame, text="Guardar Diccionario", command=self.on_ok, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief=tk.FLAT).pack(side=tk.RIGHT, padx=5)

    def on_ok(self):
        for col, obj in self.widgets.items():
            self.result[col] = {"role": obj["role"], "meaning": obj["meaning_var"].get().strip()}
        self.destroy()

# =========================================================
# APP PRINCIPAL
# =========================================================
class FinancialAssistantApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(APP_GEOMETRY)
        self.root.configure(bg="#161e2a")

        try: self.root.state('zoomed')
        except: self.root.attributes('-zoomed', True)

        self.memory = MemoryManager()
        self.df = None
        self.processor = None
        self.history = self.memory.get_recent_history() or []

        self.live_answer_started = False
        self.last_chunk = ""

        self.system_prompt_text = (
            "Eres un Científico de Datos y Analista Financiero. Tienes acceso directo a un DataFrame de pandas cargado en la variable `df`.\n\n"
            "REGLAS CRÍTICAS:\n"
            "1. REVISA SIEMPRE el diccionario de datos para saber qué significa cada columna antes de operar.\n"
            "2. Si el usuario pide un cálculo, variación o gráfico, NO intentes adivinar la respuesta. DEBES generar un bloque de código Python usando pandas/matplotlib.\n"
            "3. El código debe estar dentro de ```python y ```.\n"
            "4. Usa `print()` para mostrar resultados numéricos y `plt.show()` para gráficos.\n"
            "5. Si recibes la salida de un código previamente ejecutado, explícala en lenguaje natural sin volver a escribir código."
        )

        self.setup_ui()
        self.model_var.set(self.memory.get_last_model() or DEFAULT_MODEL)
        
        self.append_to_chat("Sistema", "📊 Asistente de Finanzas (Code Interpreter) iniciado.\nCarga un Excel para comenzar.")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        top_frame = tk.Frame(self.root, bg="#2c3e50", pady=8)
        top_frame.pack(fill=tk.X)

        lbl_title = tk.Label(top_frame, text="AInalista Financiero con Gemma 4", fg="white", bg="#2c3e50", font=("Arial", 14, "bold"))
        lbl_title.pack(side=tk.LEFT, padx=16)

        model_label = tk.Label(top_frame, text="Modelo:", fg="white", bg="#2c3e50", font=("Arial", 10, "bold"))
        model_label.pack(side=tk.LEFT, padx=(12, 6))

        self.model_var = tk.StringVar(value=DEFAULT_MODEL)
        model_entry = tk.Entry(top_frame, textvariable=self.model_var, width=18, font=("Consolas", 10))
        model_entry.pack(side=tk.LEFT, padx=(0, 10))

        self.show_thinking_var = tk.BooleanVar(value=True)
        chk_thinking = tk.Checkbutton(top_frame, text="Mostrar thinking", variable=self.show_thinking_var, bg="#2c3e50", fg="white", selectcolor="#2c3e50", activebackground="#2c3e50", activeforeground="white", font=("Arial", 10, "bold"))
        chk_thinking.pack(side=tk.LEFT, padx=(0, 12))

        btn_clear_thinking = tk.Button(top_frame, text="Limpiar thinking", command=self.clear_thinking_panel, bg="#7f8c8d", fg="white", font=("Arial", 10, "bold"), relief=tk.FLAT)
        btn_clear_thinking.pack(side=tk.LEFT, padx=(0, 12))

        btn_load = tk.Button(top_frame, text="Cargar Excel", command=self.load_excel, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief=tk.FLAT)
        btn_load.pack(side=tk.RIGHT, padx=20)

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

        thinking_title = tk.Label(thinking_header, text="Thinking Process (debug opcional)", fg="white", bg="#34495e", font=("Arial", 11, "bold"))
        thinking_title.pack(side=tk.LEFT, padx=10, pady=8)

        self.lbl_thinking_status = tk.Label(thinking_header, text="Inactivo", fg="#dfe6e9", bg="#34495e", font=("Arial", 10))
        self.lbl_thinking_status.pack(side=tk.RIGHT, padx=10)

        self.thinking_display = scrolledtext.ScrolledText(right_frame, wrap=tk.WORD, bg="#fdfdfd", font=("Consolas", 10), state=tk.DISABLED)
        self.thinking_display.pack(expand=True, fill=tk.BOTH, padx=8, pady=8)

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

    def load_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath: return
        self.btn_send.config(state=tk.DISABLED)
        self.append_to_chat("Sistema", f"📥 Cargando archivo: {os.path.basename(filepath)} ...")
        threading.Thread(target=self._load_excel_worker, args=(filepath,), daemon=True).start()

    def _load_excel_worker(self, filepath):
        try:
            excel_file = pd.ExcelFile(filepath)
            sheets = excel_file.sheet_names
            
            sheet = sheets[0]
            if len(sheets) > 1:
                dialog_result = {"sheet": None}
                def open_sheet_dialog():
                    d = SheetSelectorDialog(self.root, sheets)
                    self.root.wait_window(d)
                    dialog_result["sheet"] = d.selected_sheet
                self.root.after(0, open_sheet_dialog)
                
                import time
                while dialog_result["sheet"] is None and self.root.winfo_exists(): time.sleep(0.5)
                sheet = dialog_result["sheet"]
                if not sheet: 
                    self.root.after(0, lambda: self.btn_send.config(state=tk.NORMAL))
                    return

            temp_proc = ExcelProcessor(filepath, sheet)
            temp_proc.load()

            mem_data = self.memory.get_dataset(temp_proc.signature)
            old_defs = mem_data.get("column_definitions", {}) if mem_data else {}

            self.processor = ExcelProcessor(filepath, sheet, column_definitions=old_defs)
            self.processor.load()
            self.df = self.processor.df

            self.root.after(0, self._open_metadata_dialog)

        except Exception as e:
            err_msg = str(e) # FIX LAMBDA SCOPE
            self.root.after(0, lambda m=err_msg: self.append_to_chat("Error", f"Fallo al cargar: {m}"))
            self.root.after(0, lambda: self.btn_send.config(state=tk.NORMAL, text="Enviar"))

    def _open_metadata_dialog(self):
        dialog = ColumnDefinitionDialog(self.root, self.processor)
        self.root.wait_window(dialog)
        
        if dialog.result:
            self.processor.apply_user_definitions(dialog.result)
            payload = self.processor.build_memory_payload()
            self.memory.update_dataset(self.processor.signature, payload)
            self.append_to_chat("Sistema", "✅ Diccionario guardado.\n\n" + self.processor.build_llm_context())
        else:
            self.append_to_chat("Sistema", "Carga finalizada utilizando la autodección de la herramienta.")
            
        self.btn_send.config(state=tk.NORMAL)

    def send_message(self):
        user_text = self.input_box.get("1.0", tk.END).strip()
        if not user_text: return
        
        self.input_box.delete("1.0", tk.END)
        self.append_to_chat("Tú", user_text)
        self.btn_send.config(state=tk.DISABLED, text="Pensando...")

        self.history.append({"role": "user", "content": user_text})
        self.history = self.history[-MAX_CHAT_HISTORY:]

        threading.Thread(target=self._orchestrate_ai_workflow, daemon=True).start()

    def _orchestrate_ai_workflow(self):
        try:
            ai_reply = self._stream_ollama_call()
            self.history.append({"role": "assistant", "content": ai_reply})

            code_blocks = re.findall(r'```python(.*?)```', ai_reply, re.DOTALL)
            
            if code_blocks:
                for idx, code in enumerate(code_blocks):
                    self.root.after(0, lambda: self.append_to_chat("Sistema", f"⚙️ Ejecutando código dinámico..."))
                    
                    success, output, fig = self._run_python_sandbox(code.strip())
                    
                    if fig:
                        self.root.after(0, lambda f=fig: self._show_figure(f))
                    
                    if output.strip():
                        self.root.after(0, lambda o=output: self.append_to_chat("Consola Python", o))
                    
                    if success:
                        prompt = f"El código se ejecutó bien. Salida:\n{output}\nExplica este resultado ejecutivo."
                    else:
                        prompt = f"El código falló:\n{output}\nInforma del error y corrígelo."
                    
                    self.history.append({"role": "user", "content": prompt})
                    final_reply = self._stream_ollama_call(is_feedback=True)
                    self.history.append({"role": "assistant", "content": final_reply})

        except Exception as e:
            err_msg = str(e) # FIX LAMBDA SCOPE
            self.root.after(0, lambda m=err_msg: self.append_to_chat("Error", f"Error en el flujo: {m}"))
        finally:
            self.memory.set_recent_history(self.history)
            self.root.after(0, lambda: self.btn_send.config(state=tk.NORMAL, text="Enviar"))

    def _run_python_sandbox(self, code_string):
        if self.processor is None: return False, "Error: No hay DataFrame cargado.", None

        output_buffer = io.StringIO()
        plt.clf() 
        custom_show = {"called": False}
        def mock_show(*args, **kwargs): custom_show["called"] = True
            
        env = {'df': self.processor.df, 'pd': pd, 'plt': plt}
        env['plt'].show = mock_show

        try:
            with contextlib.redirect_stdout(output_buffer): exec(code_string, env)
            fig = plt.gcf() if (custom_show["called"] or plt.gcf().get_axes()) else None
            return True, output_buffer.getvalue(), fig
        except Exception as e:
            return False, str(e), None

    def _stream_ollama_call(self, is_feedback=False):
        model_name = self.model_var.get().strip() or DEFAULT_MODEL
        
        messages = [{"role": "system", "content": self.system_prompt_text}]
        if self.processor:
            messages.append({"role": "system", "content": "Metadatos del archivo:\n" + self.processor.build_llm_context()})

        messages.extend(self.history)

        show_think = bool(self.show_thinking_var.get())
        think_text = ""
        ans_text = ""

        self.root.after(0, self.clear_thinking_panel)
        self.root.after(0, lambda: self.set_thinking_status("Analizando..." if not is_feedback else "Concluyendo..."))
        
        sender = "Gemma" if not is_feedback else "Gemma (Análisis)"
        self.root.after(0, lambda: self._start_live_answer(sender=sender))

        try:
            stream = ollama.chat(model=model_name, messages=messages, think=show_think, stream=True, options={"temperature": 0.1})
            update_counter = 0
            
            for chunk in stream:
                msg = get_chunk_message_dict(chunk)
                
                ch_think = msg.get("thinking", "")
                if ch_think:
                    think_text += ch_think
                    if len(think_text) > 4000: think_text = think_text[-4000:]
                    if show_think and update_counter % 8 == 0:
                        self.root.after(0, lambda t=think_text: self._update_live_thinking(t))
                        self.root.after(0, lambda: self.set_thinking_status("Thinking..."))

                ch_content = msg.get("content", "")
                if ch_content and ch_content != self.last_chunk:
                    self.last_chunk = ch_content
                    ans_text += ch_content
                    self.root.after(0, lambda t=ch_content: self._append_live_answer(t))
                    
                update_counter += 1

        except Exception as e:
            err_ans = f"[Error: {e}]" # FIX LAMBDA SCOPE
            self.root.after(0, lambda m=err_ans: self._append_live_answer(m))

        self.root.after(0, lambda: self.set_thinking_status("Completado"))
        self.root.after(0, self._finish_live_answer)
        
        return ans_text

    def _show_figure(self, fig):
        win = tk.Toplevel(self.root)
        win.title("Gráfico dinámico de Pandas")
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