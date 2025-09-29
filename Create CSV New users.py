import csv
import glob
import os

from StorageFunctions import (generar_contraseña)

# ---------- Rutas de salida ----------
csv_salida = 'C:/Users/pepitoperez/Input_script.csv'
txt_salida = 'C:/Users/pepitoperez/Reporte Talento Humano.txt'

# ---------- Generador de contraseña segura ----------
generar_contraseña()

# ---------- Mapeo de departamentos a OU en AD ----------
OU_MAPPING = {
    "Administrativo": "OU=Users,OU=Administrativo,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Asesores Ventas": "OU=User,OU=Asesores Ventas,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Calidad y Logistica": "OU=Users,OU=Calidad y Logistica,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Comercial": "OU=Users,OU=Comercial,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Control de Proyectos": "OU=Users,OU=Control de Proyectos,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Diseño": "OU=Users,OU=Diseño,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Financiero": "OU=Users,OU=Financiero,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Gerencia General": "OU=Users,OU=Gerencia General,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Gerencias": "OU=User,OU=Gerencias,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Mercadeo": "OU=Users,OU=Mercadeo,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Servicio al Cliente": "OU=Users,OU=Servicio al Cliente,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Talento Humano": "OU=Users,OU=Talento Humano,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Tecnica": "OU=Users,OU=Tecnica,OU=Company,OU=CSZ,DC=cusezar,DC=corp",
    "Transformación Digital": "OU=User,OU=Transformación Digital,OU=Company,OU=CSZ,DC=cusezar,DC=corp"
}

# ---------- Buscar automáticamente el CSV más reciente ----------
carpeta_respuestas = r'C:/Users/pepitoperez/Respuestas Formulario Microsoft Forms'
archivos_csv = glob.glob(os.path.join(carpeta_respuestas, "*.csv"))

if not archivos_csv:
    raise FileNotFoundError("No se encontraron archivos CSV en la carpeta de Respuestas Formulario.")

csv_entrada = max(archivos_csv, key=os.path.getmtime)
print(f"📄 Archivo de entrada detectado automáticamente: {os.path.basename(csv_entrada)}")

usuarios_completos = []

# ---------- Leer CSV base ----------
with open(csv_entrada, mode='r', encoding='utf-8-sig') as archivo_csv:
    lector = csv.DictReader(archivo_csv, delimiter=';')
    for fila in lector:
        if all(valor.strip() == '' for valor in fila.values()):
            continue  # Saltar filas completamente vacías

        nombres = fila['Nombre(s)'].lstrip('\n').strip()
        apellidos = fila['Apellidos'].strip()
        cedula = fila['Cédula'].strip()
        oficina = fila['Oficina'].strip()
        puesto = fila['Cargo'].strip()
        departamento = fila['Departamento'].strip()
        fecha_ingreso = fila['Fecha de Ingreso'].strip()
        nombre_completo = f"{nombres} {apellidos}"

        print(f"Procesando usuario: {nombre_completo}, Fecha de ingreso: {fecha_ingreso}")
        iniciales = input("Ingresa las iniciales: ").strip().upper()
        usuario = input("Ingresa el nombre de usuario: ").strip()

        correo = f"{usuario}@cusezar.com"
        contraseña = generar_contraseña()

        ou = OU_MAPPING.get(departamento, "")
        if not ou:
            print(f"⚠️  Departamento '{departamento}' no está mapeado. Asignando OU vacío.")

        proxy = f'SMTP:{usuario}@CusezarSA.onmicrosoft.com;smtp:{correo}'

        usuario_dict = {
            "Nombres": nombres,
            "Apellidos": apellidos,
            "Iniciales": iniciales,
            "NombreCompleto": nombre_completo,
            "Oficina": oficina,
            "Correo": correo,
            "Puesto": puesto,
            "Departamento": departamento,
            "Organizacion": "Cusezar",
            "Usuario": usuario,
            "Contraseña": contraseña,
            "OU": ou,
            "proxyAddresses": proxy,
            "Cedula": cedula,
            "Fecha de Ingreso": fecha_ingreso
        }

        usuarios_completos.append(usuario_dict)

# ---------- Guardar CSV completo ----------
campos = ["Nombres", "Apellidos", "Iniciales", "NombreCompleto", "Oficina", "Correo",
          "Puesto", "Departamento", "Organizacion", "Usuario", "Contraseña", "OU", "proxyAddresses"]

with open(csv_salida, mode='w', newline='', encoding='utf-8') as archivo_csv_salida:
    escritor = csv.DictWriter(archivo_csv_salida, fieldnames=campos, delimiter=';')
    escritor.writeheader()
    for usuario in usuarios_completos:
        # Crear una copia del diccionario sin el campo 'Cedula'
        usuario_sin_cedula = {k: v for k, v in usuario.items() if k in campos}
        escritor.writerow(usuario_sin_cedula)

# ---------- Guardar TXT con formato personalizado ----------
with open(txt_salida, mode='w', encoding='utf-8') as archivo_txt:
    for usuario in usuarios_completos:
        archivo_txt.write(f"Nombre completo: {usuario['NombreCompleto']}\n")
        archivo_txt.write(f"Fecha de Ingreso: {usuario['Fecha de Ingreso']}\n")
        archivo_txt.write(f"Cédula: {usuario['Cedula']}\n")
        archivo_txt.write(f"Cargo: {usuario['Puesto']}\n")
        archivo_txt.write(f"Correo: {usuario['Correo']}\n")
        archivo_txt.write(f"Usuario: {usuario['Usuario']}\n")
        archivo_txt.write(f"Contraseña: {usuario['Contraseña']}\n\n")


print(f"Archivo CSV (input_script) y TXT (Reporte Talento Humano) generado.")
os.startfile(txt_salida)