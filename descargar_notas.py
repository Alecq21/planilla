import imaplib
import email
from email.header import decode_header
import os
import fitz  # PyMuPDF
from datetime import datetime, timedelta

# --------------------------
# CONFIGURACIÓN PERSONALIZADA
# --------------------------
USUARIO = "alequiroga1991@gmail.com"
CONTRASENA = "volj mtwp wcum qqrg"
CARPETA_BASE = "Notas"
CARPETA_VALIDACION = "Notas de Validacion"


# Lista de jurisdicciones conocidas
JURISDICCIONES = [
    "Asesoría General de Gobierno",
    "Coordinación General Unidad Gobernador",
    "Dirección General de Cultura y Educación",
    "Ministerio de Comunicación Pública",
    "Ministerio de Desarrollo Agrario",
    "Ministerio de Desarrollo de la Comunidad",
    "Ministerio de Gobierno",
    "Ministerio de Hacienda y Finanzas",
    "Ministerio de Infraestructura y Servicios Públicos",
    "Ministerio de Jefatura de Gabinete de Ministros",
    "Ministerio de Justicia y Derechos Humanos",
    "Ministerio de Mujeres y Diversidad",
    "Ministerio de Producción, Ciencia e Innovación Tecnológica",
    "Ministerio de Salud",
    "Tesorería General de la Provincia de Buenos Aires",
    "Ministerio de Seguridad",
    "Ministerio de Trabajo",
    "Secretaría General",
    "Organismo Provincial de Contrataciones",
    "Organismo Provincial de Integración Social y Urbana",
    "Ministerio de Ambiente",
    "Instituto Provincial de Lotería y Casinos",
    "Instituto de Obra Médico Asistencial",
    "Instituto de Previsión Social",
    "Instituto de la Vivienda",
    "Agencia de Recaudación de la Provincia de Buenos Aires",
    "Honorable Tribunal de Cuentas",
    "Patronato De Liberados Bonaerense",
    "Junta Electoral",
    "Contaduria General de la Provincia",
    "Agencia Administradora Estadio Único Ciudad de La Plata",
    "Organismo Provincial para el Desarrollo Sostenible",
    "Universidad Provincial de Ezeiza (UPE)",
    "Fiscalia de Estado",
    "Organismo Provincial de la Niñez y Adolescencia",
    "Ministerio de Transporte",
    "Universidad Provincial del Sudoeste (UPSO)",
    "Ministerio de Hábitat y Desarrollo Urbano",
    "Instituto Cultural",
    "Jefatura de Asesores del Gobernador",
    "Comisión de Investigaciones Científicas",
    "Comité de Cuenca del Río Reconquista",
    "Dirección de Vialidad",
    "Instituto Provincial de Asociativismo y Cooperativismo",
    "Corporación de Fomento del Valle Bonaerense del Río Colorado (CORFO)",
    "Ministerio de Economía",
    "Organismo de Control de la Energía Eléctrica de Buenos Aires",
    "Autoridad del Agua"
]

# --------------------------
# FUNCIONES
# --------------------------

def conectar_gmail():
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(USUARIO, CONTRASENA)
    return imap

def decodificar(texto):
    try:
        partes = decode_header(texto)
        resultado = ""
        for t, codificacion in partes:
            if isinstance(t, bytes):
                try:
                    resultado += t.decode(codificacion or 'utf-8', errors='replace')
                except LookupError:
                    resultado += t.decode('utf-8', errors='replace')
            else:
                resultado += t
        return resultado
    except:
        return "(asunto no legible)"

    return ''.join([str(t[0], t[1] or 'utf-8') if isinstance(t[0], bytes) else t[0] for t in partes])



def extraer_embebidos(ruta_pdf, carpeta_destino):
    doc = fitz.open(ruta_pdf)
    for i in range(doc.embfile_count()):
        info = doc.embfile_info(i)
        archivo = doc.embfile_get(i)
        nombre_archivo = info['filename']
        ruta_salida = os.path.join(carpeta_destino, nombre_archivo)
        if not os.path.exists(ruta_salida):
            with open(ruta_salida, "wb") as salida:
                salida.write(archivo)



def detectar_jurisdiccion(texto):
    for j in JURISDICCIONES:
        if j in texto:
            return j
    return "Desconocido"

def es_nota_validacion(texto):
    texto = texto.lower()
    return "secretaría general" in texto and "ministerio de economía" in texto

def generar_nombre_carpeta(base, fecha):
    carpeta = os.path.join(base, fecha)
    if not os.path.exists(carpeta):
        return carpeta
    else:
        i = 1
        while True:
            nueva = f"{carpeta}-{i}"
            if not os.path.exists(nueva):
                return nueva
            i += 1

def procesar_mails(imap):
    imap.select("inbox")

    # Filtro: últimos 1 días
    fecha_limite = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%Y')
    estado, mensajes = imap.search(None, f'(UNSEEN SINCE {fecha_limite})')

    for num in mensajes[0].split():
        _, datos = imap.fetch(num, "(RFC822)")
        mensaje = email.message_from_bytes(datos[0][1])
        asunto = decodificar(mensaje["Subject"])
        asunto_lower = asunto.lower()

        # Filtro de asunto más flexible
        if "ingresos" not in asunto_lower and "notas de ingresos" not in asunto_lower:
            continue

        print(f"📩 Procesando: {asunto}")
        if mensaje.is_multipart():
            for parte in mensaje.walk():
                if parte.get_content_disposition() == "attachment":
                    nombre = parte.get_filename()
                    if nombre and nombre.endswith(".pdf"):
                        contenido = parte.get_payload(decode=True)
                        ruta_temp = os.path.join("temp_" + nombre)
                        with open(ruta_temp, "wb") as f:
                            f.write(contenido)

                        doc = fitz.open(ruta_temp)
                        texto_pdf = ""
                        for pagina in doc:
                            texto_pdf += pagina.get_text()
                        doc.close() 

                        fecha_formato = datetime.today().strftime("%d-%m-%y")

                        if es_nota_validacion(texto_pdf):
                            carpeta_final = os.path.join(CARPETA_VALIDACION, fecha_formato)
                            os.makedirs(carpeta_final, exist_ok=True)
                            ruta_final = os.path.join(carpeta_final, nombre)
                            os.rename(ruta_temp, ruta_final)
                            print(f"✅ Nota de validación guardada en: {carpeta_final}")
                        else:
                            jurisdiccion = detectar_jurisdiccion(texto_pdf)
                            base_carpeta = os.path.join(CARPETA_BASE, jurisdiccion)
                            carpeta_final = generar_nombre_carpeta(base_carpeta, fecha_formato)
                            os.makedirs(carpeta_final, exist_ok=True)
                            ruta_final = os.path.join(carpeta_final, nombre)
                            os.rename(ruta_temp, ruta_final)
                            print(f"✅ Guardado en: {carpeta_final}")
                            extraer_embebidos(ruta_final, carpeta_final)

    imap.logout()

# --------------------------
# EJECUCIÓN PRINCIPAL
# --------------------------
# if __name__ == "__main__":
#     os.makedirs(CARPETA_BASE, exist_ok=True)
#     imap = conectar_gmail()
#     procesar_mails(imap)
if __name__ == "__main__":
    os.makedirs(CARPETA_BASE, exist_ok=True)
    os.makedirs(CARPETA_VALIDACION, exist_ok=True)
    imap = conectar_gmail()
    procesar_mails(imap)