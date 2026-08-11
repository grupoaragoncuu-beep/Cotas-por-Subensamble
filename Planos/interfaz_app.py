import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import win32com.client
import threading

import generador_vistas

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automatización de Planos - COTAS ABIGAIL")
        self.geometry("500x350")
        self.resizable(False, False)
        self.configure(padx=20, pady=20)
        
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        
        self.inv_app = None
        self.ensambles = {} 
        
        self.crear_widgets()
        self.conectar_inventor()

    def crear_widgets(self):
        ttk.Label(self, text="COTAS ABIGAIL", font=("Segoe UI", 16, "bold")).pack(pady=(0, 5))
        ttk.Label(self, text="Generador de Planos Automático", font=("Segoe UI", 10)).pack(pady=(0, 20))
        
        frame = ttk.LabelFrame(self, text=" Selecciona el Ensamble a Procesar ", padding=15)
        frame.pack(fill=tk.X, pady=10)
        
        self.combo_var = tk.StringVar()
        self.combo = ttk.Combobox(frame, textvariable=self.combo_var, state="readonly", font=("Segoe UI", 10))
        self.combo.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_refrescar = ttk.Button(frame, text="Refrescar Lista", command=self.conectar_inventor)
        self.btn_refrescar.pack(anchor=tk.E)
        
        self.btn_generar = ttk.Button(self, text="Generar Planos", command=self.iniciar_proceso)
        self.btn_generar.pack(fill=tk.X, pady=20, ipady=5)
        
        self.lbl_estado = ttk.Label(self, text="Estado: Esperando acción...", font=("Segoe UI", 9, "italic"))
        self.lbl_estado.pack(side=tk.BOTTOM, anchor=tk.W)

    def conectar_inventor(self):
        self.lbl_estado.config(text="Estado: Conectando a Inventor...")
        self.update()
        
        self.combo.set('')
        self.combo['values'] = []
        self.ensambles.clear()
        
        try:
            import inventor_com
            self.inv_app = inventor_com.conectar_inventor()
            if self.inv_app is None:
                raise Exception("No connection")
        except Exception:
            self.lbl_estado.config(text="Estado: Error - Inventor no está abierto.")
            messagebox.showerror("Error", "No se encontró Autodesk Inventor abierto.\nPor favor, abre Inventor y carga tus ensambles.")
            return

        doc_count = self.inv_app.Documents.Count
        for i in range(1, doc_count + 1):
            try:
                doc = self.inv_app.Documents.Item(i)
                if doc.DocumentType == 12291: 
                    nombre = doc.DisplayName
                    self.ensambles[nombre] = doc
            except Exception:
                pass
                
        nombres = list(self.ensambles.keys())
        if nombres:
            self.combo['values'] = nombres
            self.combo.current(0)
            self.lbl_estado.config(text=f"Estado: {len(nombres)} ensamble(s) encontrado(s).")
            self.btn_generar.config(state=tk.NORMAL)
        else:
            self.lbl_estado.config(text="Estado: No hay ensambles abiertos.")
            self.btn_generar.config(state=tk.DISABLED)
            
    def iniciar_proceso(self):
        seleccion = self.combo_var.get()
        if not seleccion or seleccion not in self.ensambles:
            messagebox.showwarning("Aviso", "Por favor selecciona un ensamble válido.")
            return
            
        self.btn_generar.config(state=tk.DISABLED)
        self.btn_refrescar.config(state=tk.DISABLED)
        self.combo.config(state=tk.DISABLED)
        self.lbl_estado.config(text="Estado: Ejecutando automatización... (Por favor espera)")
        
        t = threading.Thread(target=self.ejecutar_flujo, args=(seleccion,))
        t.start()

    def ejecutar_flujo(self, seleccion_nombre):
        import pythoncom
        pythoncom.CoInitialize()
        
        exito = False
        mensaje = ""
        try:
            import inventor_com
            inv_app_thread = inventor_com.conectar_inventor()
            if inv_app_thread is None:
                raise Exception("No se pudo conectar a Inventor en el proceso de fondo.")
                
            # Buscar el ensamble por nombre
            ensamble_doc = None
            for i in range(1, inv_app_thread.Documents.Count + 1):
                try:
                    doc = inv_app_thread.Documents.Item(i)
                    if doc.DisplayName == seleccion_nombre:
                        ensamble_doc = doc
                        break
                except:
                    pass
                    
            if ensamble_doc is None:
                raise Exception(f"No se encontró el ensamble '{seleccion_nombre}' abierto.")
                
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                
            ruta_machote = os.path.join(base_dir, "MACHOTE PLANOS.dwg")
            
            if not os.path.exists(ruta_machote):
                ruta_machote_parent = os.path.join(os.path.dirname(base_dir), "MACHOTE PLANOS.dwg")
                if os.path.exists(ruta_machote_parent):
                    ruta_machote = ruta_machote_parent
            
            if not os.path.exists(ruta_machote):
                raise FileNotFoundError(f"No se encontró el machote en:\n{ruta_machote}")
                
            self.lbl_estado.config(text="Estado: Abriendo Machote Planos...")
            machote_doc = inv_app_thread.Documents.Open(ruta_machote, True)
            
            machote_doc.Activate()
            
            self.lbl_estado.config(text="Estado: Procesando planos y cotas...")
            exito = generador_vistas.ejecutar_flujo_desde_app(inv_app_thread, ensamble_doc, machote_doc)
            
            if exito:
                mensaje = "Automatización completada exitosamente."
            else:
                mensaje = "El proceso finalizó, pero se reportaron errores en el log."
                
        except Exception as e:
            mensaje = f"Ocurrió un error inesperado:\n{str(e)}"
            exito = False
            
        finally:
            pythoncom.CoUninitialize()
            self.after(0, self.finalizar_proceso, exito, mensaje)
            
    def finalizar_proceso(self, exito, mensaje):
        self.lbl_estado.config(text="Estado: Finalizado.")
        self.btn_generar.config(state=tk.NORMAL)
        self.btn_refrescar.config(state=tk.NORMAL)
        self.combo.config(state="readonly")
        
        if exito:
            messagebox.showinfo("Completado", mensaje)
        else:
            messagebox.showerror("Error", mensaje)

if __name__ == "__main__":
    app = App()
    app.mainloop()
