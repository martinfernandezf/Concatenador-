from tkinter import Label, Button
from tkcalendar import DateEntry
from tkinter import ttk
import os
import pandas as pd
from src.procesadores.procesar import procesar_archivos
from tkinter import messagebox
from tkinter import filedialog

class VentanaPrincipal:
    def __init__(self, root):
        self.root = root
        self.root.title("Concatenador de Envíos")
        
        # Configuración de ventana
        Label(root, text="Seleccione los archivos de envíos").pack(pady=10)
        Button(root, text="Seleccionar Archivos", command=self.seleccionar_archivos).pack(pady=5)
        
        # Selección de fechas de inicio y fin
        Label(root, text="Seleccione la fecha de inicio:").pack(pady=5)
        self.fecha_inicio = DateEntry(root, width=12, background="blue", foreground="white", borderwidth=2, year=2024)
        self.fecha_inicio.pack(pady=5)
        
        Label(root, text="Seleccione la fecha de fin:").pack(pady=5)
        self.fecha_fin = DateEntry(root, width=12, background="blue", foreground="white", borderwidth=2, year=2024)
        self.fecha_fin.pack(pady=5)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        # Botón para iniciar el procesamiento
        self.button_procesar = Button(root, text="Procesar", command=self.iniciar_proceso, state="disabled")
        self.button_procesar.pack(pady=5)

        self.archivos_seleccionados = []
        
    def seleccionar_archivos(self):
        # Selección de archivos
        self.archivos_seleccionados = filedialog.askopenfilenames(title="Seleccionar archivos de Excel", filetypes=[("Archivos de Excel", "*.xlsx")])
        if self.archivos_seleccionados:
            self.button_procesar.config(state="normal")

    def iniciar_proceso(self):
        # Ejecutar el procesamiento en un hilo separado para evitar que la interfaz se congele
        if self.archivos_seleccionados:
            self.progress["value"] = 0
            self.root.update_idletasks()
            procesar_archivos(self.archivos_seleccionados, self.fecha_inicio.get(), self.fecha_fin.get(), self.progress, self.root)
        else:
            messagebox.showwarning("Advertencia", "Debe seleccionar archivos primero.")
