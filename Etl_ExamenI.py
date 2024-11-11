import pandas as pd
import pyodbc
from tkinter import filedialog, messagebox
import tkinter as tk

def crear_tablas(cursor):
    # SQL para crear las tablas si no existen
    cursor.execute('''
    IF OBJECT_ID('Pacientes', 'U') IS NULL
    CREATE TABLE Pacientes (
        PacienteID INT PRIMARY KEY,
        Nombre NVARCHAR(50),
        Edad INT,
        Genero NVARCHAR(10),
        Direccion NVARCHAR(100)
    )
    ''')

    cursor.execute('''
    IF OBJECT_ID('Enfermeras', 'U') IS NULL
    CREATE TABLE Enfermeras (
        EnfermeraID INT PRIMARY KEY,
        Nombre NVARCHAR(50),
        Turno NVARCHAR(20),
        Especialidad NVARCHAR(50)
    )
    ''')

    cursor.execute('''
    IF OBJECT_ID('Consultas', 'U') IS NULL
    CREATE TABLE Consultas (
        ConsultaID INT PRIMARY KEY,
        PacienteID INT FOREIGN KEY REFERENCES Pacientes(PacienteID),
        EnfermeraID INT FOREIGN KEY REFERENCES Enfermeras(EnfermeraID),
        Fecha DATE,
        Diagnostico NVARCHAR(255)
    )
    ''')

    cursor.execute('''
    IF OBJECT_ID('Tratamientos', 'U') IS NULL
    CREATE TABLE Tratamientos (
        TratamientoID INT PRIMARY KEY,
        ConsultaID INT FOREIGN KEY REFERENCES Consultas(ConsultaID),
        Medicamento NVARCHAR(50),
        Duracion INT,  -- Duracion es un entero en días
        Dosis NVARCHAR(50)
    )
    ''')

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

            # Limpiar la columna 'Duracion' en df_tratamientos
            # Extraer solo los números de Duracion; si no hay número, poner NaN
            df_tratamientos['Duracion'] = pd.to_numeric(df_tratamientos['Duracion'].str.extract(r'(\d+)')[0], errors='coerce').astype('Int64')

            
            # Configurar la conexión a SQL Server
            conexion = pyodbc.connect(
                'DRIVER={SQL Server};'
                'SERVER=DESKTOP-RCPI344;'
                'DATABASE=ExamenEtl;'
                'Trusted_Connection=yes;'
            )
            cursor = conexion.cursor()

            # Crear tablas si no existen
            crear_tablas(cursor)
            conexion.commit()

            # Insertar datos en la tabla Pacientes
            for _, row in df_pacientes.iterrows():
                cursor.execute('''
                    INSERT INTO Pacientes (PacienteID, Nombre, Edad, Genero, Direccion)
                    VALUES (?, ?, ?, ?, ?)
                ''', row['PacienteID'], row['Nombre'], row['Edad'], row['Genero'], row['Direccion'])
            
            # Insertar datos en la tabla Enfermeras
            for _, row in df_enfermeras.iterrows():
                cursor.execute('''
                    INSERT INTO Enfermeras (EnfermeraID, Nombre, Turno, Especialidad)
                    VALUES (?, ?, ?, ?)
                ''', row['EnfermeraID'], row['Nombre'], row['Turno'], row['Especialidad'])
            
            # Insertar datos en la tabla Consultas
            for _, row in df_consultas.iterrows():
                cursor.execute('''
                    INSERT INTO Consultas (ConsultaID, PacienteID, EnfermeraID, Fecha, Diagnostico)
                    VALUES (?, ?, ?, ?, ?)
                ''', row['ConsultaID'], row['PacienteID'], row['EnfermeraID'], row['Fecha'], row['Diagnostico'])
            
            # Insertar datos en la tabla Tratamientos
            for _, row in df_tratamientos.iterrows():
                cursor.execute('''
                    INSERT INTO Tratamientos (TratamientoID, ConsultaID, Medicamento, Duracion, Dosis)
                    VALUES (?, ?, ?, ?, ?)
                ''', row['TratamientoID'], row['ConsultaID'], row['Medicamento'], row['Duracion'] if pd.notna(row['Duracion']) else None, row['Dosis'])

            # Confirmar los cambios en la base de datos
            conexion.commit()
            messagebox.showinfo("Éxito", "Datos cargados exitosamente en SQL Server.")

        except Exception as e:
            # Mostrar el error en caso de que ocurra
            messagebox.showerror("Error", f"Ocurrió un error: {e}")
        
        finally:
            # Cerrar cursor y conexión
            cursor.close()
            conexion.close()
            root.close()

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("ETL Sistema de Enfermería")
root.geometry("400x200")

# Botón para cargar el archivo Excel y ejecutar el proceso ETL
btn_cargar = tk.Button(root, text="Cargar Archivo Excel y Ejecutar ETL", command=cargar_datos)
btn_cargar.pack(pady=50)

root.mainloop()
