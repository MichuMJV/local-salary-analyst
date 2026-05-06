import pandas as pd
import ollama
import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading
import os

# CONFIGURACIÓN
# Asegúrate de tener tu modelo de Gemma corriendo en Ollama localmente
MODELO_AI = "gemma4" 

class FinancialAssistantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gemma - Analista de Finanzas")
        self.root.geometry("900x700")
        self.root.configure(bg="#161e2a")
        
        self.df = None
        self.excel_context = ""
        
        # Prompt base
        self.system_prompt = {
            "role": "system",
            "content": """Eres un experto analista financiero 
Tu objetivo es resolver problemas de análisis de datos, cierres y automatizaciones.
No inventes datos. Basate estrictamente en la información que se te proporcione."""
        }
        
        self.history = [] 
        self.setup_ui()
        self.append_to_chat("Sistema", "📊 Asistente iniciado. Listo para analizar. Carga tu Excel para comenzar.")

    def setup_ui(self):
        top_frame = tk.Frame(self.root, bg="#2c3e50", pady=10)
        top_frame.pack(fill=tk.X)
        
        lbl_title = tk.Label(top_frame, text="AInalista Financiero", fg="white", bg="#2c3e50", font=("Arial", 14, "bold"))
        lbl_title.pack(side=tk.LEFT, padx=20)
        
        btn_load = tk.Button(top_frame, text="Cargar Excel Completo", command=self.load_excel, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), relief=tk.FLAT)
        btn_load.pack(side=tk.RIGHT, padx=20)
        
        self.chat_display = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, bg="white", font=("Consolas", 11), state=tk.DISABLED)
        self.chat_display.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        
        bottom_frame = tk.Frame(self.root, bg="#2980b9")
        bottom_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        self.input_box = tk.Text(bottom_frame, height=4, font=("Consolas", 11))
        self.input_box.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(0, 10))
        self.input_box.bind("<Return>", self.handle_return)
        
        self.btn_send = tk.Button(bottom_frame, text="Enviar", command=self.send_message, bg="#2980b9", fg="white", font=("Arial", 11, "bold"), width=10, relief=tk.FLAT)
        self.btn_send.pack(side=tk.RIGHT, fill=tk.Y)

    def load_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
            
        try:
            self.append_to_chat("Sistema", f"Cargando y procesando archivo completo: {os.path.basename(filepath)}...")
            
            # Leemos el Excel completo
            self.df = pd.read_excel(filepath).fillna("")
            
            # Transformamos TODO el dataframe a formato CSV delimitado por barras (|) 
            # Esto es muy amigable para que el LLM lo lea como una tabla
            datos_completos = self.df.to_csv(index=False, sep='|')
            
            # Guardamos el Excel completo en la variable de contexto
            self.excel_context = f"""
--- INICIO DE LOS DATOS COMPLETOS DEL EXCEL ---
{datos_completos}
--- FIN DE LOS DATOS COMPLETOS DEL EXCEL ---
Utiliza esta data exacta y completa para tus cálculos y respuestas.
"""
            # Advertencia de tamaño
            filas, columnas = self.df.shape
            self.append_to_chat("Sistema", f"✅ Excel memorizado al 100%. Se cargaron {filas} filas y {columnas} columnas. El contenido completo volará🛩️ en cada pregunta.")
            
        except Exception as e:
            self.append_to_chat("Error", f"No se pudo cargar el Excel: {str(e)}")

    def handle_return(self, event):
        if not event.state & 0x0001: 
            self.send_message()
            return "break"

    def send_message(self):
        user_text = self.input_box.get("1.0", tk.END).strip()
        if not user_text:
            return
            
        self.input_box.delete("1.0", tk.END)
        self.append_to_chat("Tú", user_text)
        self.btn_send.config(state=tk.DISABLED, text="Pensando...")
        
        threading.Thread(target=self._process_ai_response, args=(user_text,), daemon=True).start()

    def _process_ai_response(self, user_text):
        try:
            self.history.append({"role": "user", "content": user_text})
            
            messages_to_send = [self.system_prompt] + self.history[:-1]
            
            # INYECCIÓN DEL EXCEL COMPLETO
            # Insertamos los datos completos justo antes de enviar tu pregunta actual
            if self.excel_context:
                messages_to_send.append({
                    "role": "system",
                    "content": self.excel_context
                })
                
            messages_to_send.append(self.history[-1])
            
            response = ollama.chat(model=MODELO_AI, messages=messages_to_send)
            ai_reply = response['message']['content']
            
            self.history.append({"role": "assistant", "content": ai_reply})
            self.root.after(0, self.append_to_chat, "Gemma", ai_reply)
            
        except Exception as e:
            error_msg = f"Error conectando con la IA: {str(e)}"
            self.root.after(0, self.append_to_chat, "Error", error_msg)
        finally:
            self.root.after(0, lambda: self.btn_send.config(state=tk.NORMAL, text="Enviar"))

    def append_to_chat(self, sender, text):
        self.chat_display.config(state=tk.NORMAL)
        
        colores = {"Tú": "#2980b9", "Gemma": "#27ae60", "Sistema": "#8e44ad", "Error": "#c0392b"}
        color = colores.get(sender, "#000000")
            
        self.chat_display.tag_config(sender, foreground=color, font=("Consolas", 11, "bold"))
        self.chat_display.insert(tk.END, f"{sender}:\n", sender)
        self.chat_display.insert(tk.END, f"{text}\n\n")
        
        self.chat_display.yview(tk.END)
        self.chat_display.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = FinancialAssistantApp(root)
    root.mainloop()