import os
from tkinter import Tk, filedialog, messagebox
from gui.interfaz import VentanaPrincipal

# Crear la ventana principal
if __name__ == "__main__":
    root = Tk()
    app = VentanaPrincipal(root)
    root.mainloop()
