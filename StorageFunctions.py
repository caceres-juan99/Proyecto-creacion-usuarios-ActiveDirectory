import pygetwindow as gw
import subprocess
import threading
import random
import string

ruta_txt = r"C:\Users\pepitoperez\Reporte Talento Humano.txt"

# ---------- Generador de contraseña segura ----------
def generar_contraseña(longitud=14):
    
    # ---------- Configuración de símbolos permitidos ----------
    SIMBOLOS_EXCLUIDOS = ":;^~'`\\\""
    SIMBOLOS_PERMITIDOS = ''.join(c for c in string.punctuation if c not in SIMBOLOS_EXCLUIDOS)

    if longitud < 4:
        raise ValueError("La contraseña debe tener al menos 4 caracteres.")
    mayuscula = random.choice(string.ascii_uppercase)
    minuscula = random.choice(string.ascii_lowercase)
    numero = random.choice(string.digits)
    simbolo = random.choice(SIMBOLOS_PERMITIDOS)
    caracteres_restantes = longitud - 4
    todos_los_caracteres = string.ascii_letters + string.digits + SIMBOLOS_PERMITIDOS
    relleno = [random.choice(todos_los_caracteres) for _ in range(caracteres_restantes)]
    contraseña_lista = list(mayuscula + minuscula + numero + simbolo) + relleno
    random.shuffle(contraseña_lista)
    return ''.join(contraseña_lista)

# ---------- Funciones para los botones ----------
def CSVgenerator_action():
    def task():
        # Cambiar ruta de acceso de acuerdo a la ubicación del archivo "Create CSV New users.py"
        subprocess.run(["python", "C:/Users/pepitoperez/Create CSV New users.py"])
    threading.Thread(target=task).start()

def ScriptExecution_action():
    # Ruta de tu script PowerShell
    script_ps = r"C:\Users\pepitoperez\SCRIPT - Crear usuarios nuevos en AD.ps1"
    
    # Ejecutar el script en PowerShell 7
    subprocess.run([
        r"C:\Program Files\PowerShell\7\pwsh.exe",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", script_ps
    ])

def leer_usuarios_desde_txt(ruta_archivo):
    with open(ruta_archivo, "r", encoding="utf-8") as f:
        contenido = f.read().strip()

    bloques = contenido.split("\n\n")  # separa usuarios por línea en blanco
    usuarios = []
    
    for bloque in bloques:
        datos = {}
        for linea in bloque.split("\n"):
            if ": " in linea:
                clave, valor = linea.split(": ", 1)
                datos[clave.strip()] = valor.strip()
        usuarios.append(datos)
    return usuarios