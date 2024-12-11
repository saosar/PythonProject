"""
Script para automatización en MS Project utilizando Python
"""

# Importar librerías necesarias
from win32com.client import Dispatch
import os

# Paso 1: Ruta de los archivos
ruta_principal = r"C:/Users/sarao/Downloads/INTERVENTORIA/CronogramaV01/OptimizaInforme"
archivo_project = "Proyecto2015.mpp"  # Archivo de MS Project

# Paso 2: Conectar al archivo de MS Project
ms_project = Dispatch("MSProject.Application")
ms_project.Visible = True  # Abre MS Project para visualizar los cambios

# Abre el archivo .mpp
ruta_archivo = os.path.join(ruta_principal, archivo_project)
ms_project.FileOpen(ruta_archivo)

# Obtén la instancia del proyecto activo
proyecto = ms_project.ActiveProject

# Paso 3: Crear tareas en el archivo de MS Project
tareas = proyecto.Tasks

# Añadir tareas al proyecto (ejemplo)
tareas.Add("Tarea 1: Configuración inicial")
tareas.Add("Tarea 2: Validación del diseño")
tareas.Add("Tarea 3: Revisión de cálculos y parámetros")
tareas.Add("Tarea 4: Ejecución de simulaciones")
tareas.Add("Tarea 5: Revisión y análisis final")
tareas.Add("Tarea 6: Generación del informe ejecutivo")

# Paso 4: Guardar y cerrar
ms_project.FileSave()
ms_project.Quit()

print("Tareas añadidas al archivo de MS Project y guardadas correctamente.")
