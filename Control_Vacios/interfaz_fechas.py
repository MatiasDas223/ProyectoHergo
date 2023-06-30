# Archivo: interfaz_fechas.py

import tkinter as tk
from tkinter import messagebox
import datetime

class VentanaFechas(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Selección de Fechas")

        self.fecha_inicio = datetime.date.today() - datetime.timedelta(days=14)
        self.fecha_fin = datetime.date.today() - datetime.timedelta(days=7)

        # Etiqueta y campo de entrada para la fecha de inicio
        label_inicio = tk.Label(self, text="Fecha de inicio:")
        label_inicio.pack()
        self.entry_inicio = tk.Entry(self)
        self.entry_inicio.pack()

        # Etiqueta y campo de entrada para la fecha de fin
        label_fin = tk.Label(self, text="Fecha de fin:")
        label_fin.pack()
        self.entry_fin = tk.Entry(self)
        self.entry_fin.pack()

        # Botón para obtener las fechas ingresadas
        boton_obtener_fechas = tk.Button(self, text="Obtener Fechas", command=self.obtener)
        boton_obtener_fechas.pack()

    def obtener(self):
        fecha_inicio = self.entry_inicio.get()
        fecha_fin = self.entry_fin.get()

        try:
            self.fecha_inicio = datetime.datetime.strptime(fecha_inicio, "%d/%m/%Y").date()
            self.fecha_fin = datetime.datetime.strptime(fecha_fin, "%d/%m/%Y").date()
            self.destroy()  # Cerrar la ventana
        except ValueError:
            messagebox.showerror("Error", "Las fechas ingresadas no son válidas")

def obtener_fechas():
    ventana = tk.Tk()
    ventana.withdraw()  # Ocultar la ventana principal

    ventana_fechas = VentanaFechas(ventana)
    ventana_fechas.wait_window(ventana_fechas)

    fecha_inicio = ventana_fechas.fecha_inicio
    fecha_fin = ventana_fechas.fecha_fin

    if fecha_inicio and fecha_fin:
        return fecha_inicio, fecha_fin

    return None
