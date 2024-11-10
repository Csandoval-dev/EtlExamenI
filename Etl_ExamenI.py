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
    if archivo_excel:
         try:
            # Cargar cada hoja del archivo Excel en DataFrames
            df_pacientes = pd.read_excel(archivo_excel, sheet_name='Pacientes')
            df_enfermeras = pd.read_excel(archivo_excel, sheet_name='Enfermeras')
            df_consultas = pd.read_excel(archivo_excel, sheet_name='Consultas')
            df_tratamientos = pd.read_excel(archivo_excel, sheet_name='Tratamientos')
            
            # Configurar la conexión a SQL Server
            conexion = pyodbc.connect(
                'DRIVER={SQL Server};'
                'SERVER=DESKTOP-RCPI344;'
                'DATABASE=ExamenEtl;'
            )
            cursor = conexion.cursor()