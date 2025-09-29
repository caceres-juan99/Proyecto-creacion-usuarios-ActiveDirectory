import pyautogui
import tkinter as tk
from tkinter import PhotoImage
from PIL import Image, ImageTk
import pygetwindow as gw
import sys

# Ruta completa donde está StorageFunctions.py
ruta_storage = r"C:\Users\jcaceres\OneDrive - Cusezar S.A\Cusezar S.A\SCRIPTs\PYTHON - Crear usuarios nuevos en AD\Create CSV New users - Power Automate"
if ruta_storage not in sys.path:
    sys.path.append(ruta_storage)

from StorageFunctions import (CSVgenerator_action, ScriptExecution_action)

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Nuevos Ingresos Cusezar")

# Establecer el icono de la ventana
# Cambiar ruta de acceso de acuerdo a la ubicación del icono "Icono de ventana.ico"
ventana.iconbitmap("C:/Users/jcaceres/OneDrive - Cusezar S.A/Cusezar S.A/SCRIPTs/PYTHON - Crear usuarios nuevos en AD/Create CSV New users - Power Automate/Interfaz de usuario/Iconos/Icono de ventana.ico")

# Ajustar el tamaño de la ventana
ventana.geometry("400x400")
ventana.resizable(False, False)  # Evitar que la ventana cambie de tamaño

# Cargar la imagen de fondo
# Cambiar ruta de acceso de acuerdo a la ubicación del fondo "Fondo de interfaz.py"
fondo_image = Image.open("C:/Users/jcaceres/OneDrive - Cusezar S.A/Cusezar S.A/SCRIPTs/PYTHON - Crear usuarios nuevos en AD/Create CSV New users - Power Automate/Interfaz de usuario/Iconos/Fondo de interfaz.png")
fondo_image = fondo_image.resize((400, 400), Image.Resampling.LANCZOS)
fondo_photo = ImageTk.PhotoImage(fondo_image)

# Crear un widget de etiqueta para mostrar la imagen de fondo
label_fondo = tk.Label(ventana, image=fondo_photo)
label_fondo.place(x=0, y=0, relwidth=1, relheight=1)

# Crear el logo de la compañía con subtítulo
logo_frame = tk.Frame(ventana, bg="white")

logo_frame.pack(pady=30)

# Crear un frame transparente para los botones
frame_botones = tk.Frame(ventana, bg='', relief='flat')
frame_botones.pack(pady=100)

# Crear botones con estilo
boton_CSVgenerator = tk.Button(
    frame_botones,
    text="CSV Generator",
    command=CSVgenerator_action,
    bg='#dfdfdf', 
    fg='black',
    font=("Lato", 10, "bold"),
    width=14,
    height=1,
    relief="solid"
)

# Crear botones con estilo
boton_ScriptExecution = tk.Button(
    frame_botones,
    text="New user AD",
    command=ScriptExecution_action,
    bg='#dfdfdf', 
    fg='black',
    font=("Lato", 10, "bold"),
    width=14,
    height=1,
    relief="solid"
)

# Colocar botones en el frame

boton_CSVgenerator.pack(pady=5)
boton_ScriptExecution.pack(pady=5)

# Iniciar el bucle principal de Tkinter
ventana.mainloop()