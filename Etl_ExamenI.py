import pandas as pd
import pyodbc
from tkinter import filedialog, messagebox
import tkinter as tk

def cargar_datos():
    # Abrir cuadro de diálogo para seleccionar el archivo Excel
    archivo_excel = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )