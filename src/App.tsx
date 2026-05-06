import { useState } from 'react';
import { motion } from 'motion/react';
import { 
  FileCode, 
  Terminal, 
  ShieldCheck, 
  Download, 
  ClipboardCheck, 
  Database,
  Cpu,
  Info,
  Lock,
  ArrowRight,
  Settings
} from 'lucide-react';

export default function App() {
  const [copied, setCopied] = useState(false);

  const pythonCode = `import pandas as pd
import ollama
import sys
import os

# Asegúrate de haber instalado: pip install pandas openpyxl ollama
MODELO_AI = "gemma:2b" 

def analizar_excel(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo)
        columnas = df.columns.tolist()
        resumen_datos = df.head(3).to_string()
        
        contexto = f"Columnas: {columnas}\\nMuestra:\\n{resumen_datos}"
        return df, contexto
    except Exception as e:
        print(f"Error: {e}")
        return None, None

# ... (Script completo disponible en el proyecto)`;

  const copyToClipboard = () => {
    navigator.clipboard.writeText(pythonCode);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="min-h-screen geometric-bg text-slate-800 font-sans selection:bg-indigo-100 selection:text-indigo-900">
      
      {/* Navigation Bar */}
      <nav className="h-16 bg-slate-900 text-white flex items-center justify-between px-6 border-b border-slate-700 shadow-lg sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-indigo-500 rounded flex items-center justify-center font-bold text-white shadow-sm ring-1 ring-white/20">
            G
          </div>
          <span className="text-lg font-semibold tracking-tight hidden sm:inline">
            Local Salary Analyst <span className="text-indigo-400 font-mono text-sm ml-2">v1.0</span>
          </span>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 px-3 py-1 bg-emerald-500/10 border border-emerald-500/30 rounded-full text-[11px] text-emerald-400 font-medium">
            <div className="status-dot bg-emerald-500 animate-pulse"></div>
            MODO LOCAL (100% OFFLINE)
          </div>
          <button className="p-2 hover:bg-slate-800 rounded-lg transition-colors text-slate-400">
            <Settings className="w-4 h-4" />
          </button>
        </div>
      </nav>

      {/* Main Content Area */}
      <main className="max-w-[1280px] mx-auto p-6 md:p-8 grid grid-cols-1 lg:grid-cols-12 gap-8">
        
        {/* Left Sidebar - Meta Info & Stats */}
        <aside className="lg:col-span-3 space-y-6">
          <div className="glass p-6 rounded-2xl shadow-sm border border-slate-200 flex flex-col gap-5">
            <h2 className="text-[10px] uppercase font-bold tracking-[0.2em] text-slate-400">Estado de Seguridad</h2>
            <div className="space-y-4">
              <div className="flex items-center gap-3 p-3 bg-white/50 rounded-xl border border-slate-100">
                <Lock className="w-5 h-5 text-emerald-600" />
                <div>
                  <p className="text-xs font-bold text-slate-900">Cifrado Local</p>
                  <p className="text-[10px] text-slate-500 uppercase">Sin conexión externa</p>
                </div>
              </div>
              <div className="flex items-center gap-3 p-3 bg-white/50 rounded-xl border border-slate-100">
                <ShieldCheck className="w-5 h-5 text-indigo-600" />
                <div>
                  <p className="text-xs font-bold text-slate-900">Privacidad Total</p>
                  <p className="text-[10px] text-slate-500 uppercase">Datos no subidos</p>
                </div>
              </div>
            </div>
          </div>

          <div className="glass p-6 rounded-2xl shadow-sm border border-slate-200">
            <h2 className="text-[10px] uppercase font-bold tracking-[0.2em] text-slate-400 mb-6">Motor IA Recomendado</h2>
            <div className="space-y-5 text-sm">
              <div className="flex justify-between items-center text-xs">
                <span className="text-slate-500">Arquitectura:</span>
                <span className="font-mono font-bold text-slate-900">Gemma-2b-IT</span>
              </div>
              <div className="flex justify-between items-center text-xs">
                <span className="text-slate-500">Recurso RAM:</span>
                <span className="font-mono font-bold text-indigo-600">~2.5 GB</span>
              </div>
              <div className="pt-4 border-t border-slate-200">
                <p className="text-[10px] font-bold text-slate-400 uppercase mb-3 leading-tight">Optimizado Para</p>
                <div className="flex flex-wrap gap-2">
                  <span className="px-2 py-1 bg-slate-200 rounded text-[9px] font-bold">EXCEL</span>
                  <span className="px-2 py-1 bg-slate-200 rounded text-[9px] font-bold">SALARIOS</span>
                  <span className="px-2 py-1 bg-slate-200 rounded text-[9px] font-bold">PRIVACIDAD</span>
                </div>
              </div>
            </div>
          </div>
        </aside>

        {/* Center/Right Main Section */}
        <div className="lg:col-span-9 space-y-8">
          
          {/* Hero Welcome Section */}
          <div className="glass rounded-3xl p-8 border border-white/50 shadow-sm relative overflow-hidden">
             <div className="absolute top-0 right-0 p-8 text-indigo-50/20 translate-x-12 -translate-y-12">
               <Cpu className="w-64 h-64 rotate-12" />
             </div>
             <div className="relative z-10">
               <h1 className="text-4xl font-bold tracking-tight text-slate-900 mb-3">Analizador de Nómina Local</h1>
               <p className="text-slate-500 text-lg max-w-2xl leading-relaxed">
                 Ejecuta un cerebro de inteligencia artificial directamente en tu laptop. 
                 Sin internet, sin pagos, sin riesgo de filtración de datos sensibles.
               </p>
             </div>
          </div>

          {/* Requirements Grid (Geometric style) */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {[
              { icon: Terminal, title: "Python 3.10+", desc: "El lenguaje base", bg: "bg-indigo-50", color: "text-indigo-600" },
              { icon: Cpu, title: "Ollama", desc: "El motor de modelos", bg: "bg-emerald-50", color: "text-emerald-600" },
              { icon: Database, title: "Bibliotecas", desc: "Pandas & OpenPyXL", bg: "bg-purple-50", color: "text-purple-600" }
            ].map((req, i) => (
              <motion.div 
                key={i}
                whileHover={{ y: -4, scale: 1.02 }}
                className="bg-white p-6 rounded-2xl shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-slate-100 flex flex-col items-center text-center"
              >
                <div className={`w-12 h-12 ${req.bg} ${req.color} rounded-2xl flex items-center justify-center mb-4`}>
                  <req.icon className="w-6 h-6" />
                </div>
                <h3 className="font-bold text-slate-900">{req.title}</h3>
                <p className="text-slate-500 text-xs mt-1">{req.desc}</p>
              </motion.div>
            ))}
          </div>

          {/* Implementation Steps Section */}
          <section className="bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="bg-slate-900 px-8 py-5 flex items-center justify-between border-b border-slate-800">
              <div className="flex items-center gap-3">
                <FileCode className="w-5 h-5 text-indigo-400" />
                <h2 className="text-white font-semibold text-sm tracking-wide uppercase">Guía de Implementación</h2>
              </div>
              <div className="flex gap-2">
                <div className="w-2 h-2 rounded-full bg-slate-700"></div>
                <div className="w-2 h-2 rounded-full bg-slate-700"></div>
                <div className="w-2 h-2 rounded-full bg-slate-700"></div>
              </div>
            </div>
            
            <div className="p-8">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-start">
                <div className="space-y-6">
                  <h3 className="font-bold text-slate-900 flex items-center gap-2 underline decoration-indigo-200 underline-offset-4">
                    <Info className="w-4 h-4 text-indigo-500" />
                    Pasos para el éxito
                  </h3>
                  <div className="space-y-6">
                    {[
                      "Descarga 'analista_local.py' del explorador de archivos.",
                      "Asegúrate de tener Ollama corriendo y el modelo bajado.",
                      "Lanza el script con 'python analista_local.py'.",
                      "Arrastra tu Excel y empieza el análisis privado."
                    ].map((step, idx) => (
                      <div key={idx} className="flex gap-4 group">
                        <span className="flex-none w-6 h-6 rounded-full bg-slate-100 flex items-center justify-center text-[10px] font-bold text-slate-400 group-hover:bg-indigo-500 group-hover:text-white transition-colors duration-300">
                          {idx + 1}
                        </span>
                        <p className="text-sm text-slate-600 leading-snug">{step}</p>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="space-y-4">
                   <div className="flex items-center justify-between mb-2">
                    <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Vista Previa del Script</span>
                    <button 
                      onClick={copyToClipboard}
                      className="flex items-center gap-2 text-xs bg-slate-900 text-white px-4 py-1.5 rounded-full hover:bg-black transition-all active:scale-95 shadow-lg shadow-slate-200"
                    >
                      {copied ? <ClipboardCheck className="w-3.5 h-3.5" /> : <Download className="w-3.5 h-3.5 text-indigo-400" />}
                      {copied ? "¡Copiado!" : "Copiar SDK"}
                    </button>
                  </div>
                  <div className="relative group">
                    <div className="absolute -inset-1 bg-gradient-to-r from-indigo-500 to-emerald-500 rounded-2xl blur opacity-20 group-hover:opacity-40 transition duration-1000 group-hover:duration-200"></div>
                    <pre className="relative bg-slate-900 text-indigo-300/80 p-6 rounded-2xl overflow-x-auto text-[11px] font-mono leading-relaxed border border-slate-800 shadow-2xl">
                      {pythonCode}
                      {"\n\n# Consulta el archivo /analista_local.py"}
                      {"\n# para ver la lógica completa de chat."}
                    </pre>
                  </div>
                </div>
              </div>
            </div>
          </section>

          {/* Footer Card */}
          <div className="h-32 glass rounded-3xl border border-slate-200 flex flex-wrap items-center px-10 gap-x-12 py-6">
            <div className="flex flex-col">
              <p className="text-3xl font-black text-slate-900 leading-tight">0%</p>
              <p className="text-[9px] text-slate-400 uppercase tracking-widest font-bold">Fuga de Datos</p>
            </div>
            <div className="hidden md:block h-12 w-px bg-slate-200"></div>
            <div className="flex flex-col">
              <p className="text-3xl font-black text-slate-900 leading-tight">Gemma</p>
              <p className="text-[9px] text-slate-400 uppercase tracking-widest font-bold">Motor Recomendado</p>
            </div>
            <div className="hidden md:block h-12 w-px bg-slate-200"></div>
            <div className="flex-1 min-w-[200px]">
              <div className="flex justify-between items-center mb-2">
                <span className="text-[9px] font-bold text-slate-500 uppercase">Privacidad en Ejecución</span>
                <span className="text-[10px] font-bold text-indigo-600 px-2 py-0.5 bg-indigo-50 rounded-full">ESTRICTA</span>
              </div>
              <div className="w-full h-2 by-slate-200 bg-slate-100/50 rounded-full flex p-0.5 overflow-hidden ring-1 ring-slate-200">
                <div className="bg-gradient-to-r from-indigo-500 to-indigo-400 h-full w-[100%] rounded-full shadow-[0_0_8px_rgba(79,70,229,0.4)]"></div>
              </div>
            </div>
          </div>
        </div>

      </main>

      {/* Persistent Footer Bar */}
      <footer className="h-10 bg-white border-t border-slate-200 flex items-center justify-between px-8 text-[10px] text-slate-400 font-bold uppercase tracking-widest">
        <div className="flex gap-6">
          <span className="flex items-center gap-1.5"><ArrowRight className="w-3 h-3" /> Entorno: Laptop Local</span>
          <span className="flex items-center gap-1.5"><ArrowRight className="w-3 h-3" /> Datos: Nómina Privada</span>
        </div>
        <div>
          &copy; 2026 GEMMA ENGINE LABS · SIN ACCESO A NUBE
        </div>
      </footer>
    </div>
  );
}
