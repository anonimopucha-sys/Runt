
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import datetime
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as messagebox
from tkinter import PhotoImage
from dateutil.parser import parse
import os
import re
import time

import win32com.client as win32  # Outlook COM

# =========================================================
# CONFIGURACIÓN
# =========================================================
BASE_IMG_DIR   = r"C:\RPA\CE0864_Consulta_Runt_Simit\Aut. Runt x Cedula y Placa\ImgText"
PLANTILLA_SOAT = os.path.join(BASE_IMG_DIR, "soat.jpg")
PLANTILLA_RTM  = os.path.join(BASE_IMG_DIR, "rtm.jpg")

SALIDA_SOAT = os.path.join(BASE_IMG_DIR, "New Imagen 1.jpg")
SALIDA_RTM  = os.path.join(BASE_IMG_DIR, "New Imagen 2.jpg")

# ⚠ Revisa ruta y nombre exacto del Excel
EXCEL_PATH  = r"C:\CE0864_Consulta_Runt_Simit\Consulta_RUNT_VIM.xlsx"
EXCEL_SHEET = "Listado VIM"  # leer explícitamente Listado VIM

# Columnas EXACTAS (Listado VIM)
CEDULA_COL            = 'Numero de Cedula'
PLACA_COLUMN          = 'Placa'
ESTADO_SOAT_COL       = 'Estado del SOAT'            # SOAT: notificar si == 'NO VIGENTE' Y fecha < hoy
FECHA_SOAT_COL        = 'Vigencia del SOAT hasta'    # fecha que se usa para la métrica
VIGENCIA_RTM_COL      = 'Vigencia RTM'               # RTM : notificar si == 'NO' Y fecha < hoy (o 'NO VIGENTE')
FECHA_RTM_COL         = 'Fecha vigencia RTM'         # fecha que se usa para la métrica
FECHA_EXPEDICION_RTM  = 'Fecha expedicion RTM'       # solo informativa
CORREOS_COL           = 'Correos'

# (Opcional) Si quieres copia oculta a alguien, pon aquí el correo. Si no, deja lista vacía:
DEFAULT_BCC = []  # ejemplo: ["servicioalempleado@cens.com.co"]

# =========================================================
# UTILIDADES TEXTO/FECHA
# =========================================================
def norm_text_exact(value: str) -> str:
    """
    Normalización robusta para comparación EXACTA:
      - Convierte a str
      - Remueve NBSP (\xa0), BOM (\ufeff), zero‑width (\u200b)
      - Trim
      - Colapsa múltiples espacios a uno
      - Convierte a MAYÚSCULAS
    """
    s = str(value or "")
    s = s.replace('\xa0', ' ').replace('\ufeff', '').replace('\u200b', '')
    s = s.strip()
    s = re.sub(r'\s+', ' ', s)
    return s.upper()

def parse_mixed_date(value):
    """Parser robusto para dd/mm/yyyy o mm/dd/yyyy; ignora 'Sin info'/'Error'. Se usa para la métrica y para mostrar."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.date() if isinstance(value, datetime.datetime) else value

    s = str(value).strip()
    s_lower = s.lower()
    if not s or 'sin info' in s_lower or 'error' in s_lower:
        return None

    s_norm = re.sub(r'[\.\-\s]', '/', s)
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{2,4})$', s_norm)
    if m:
        a, b, c = m.groups()
        a = int(a); b = int(b); c = int(c)
        if c < 100: c += 2000

        # Heurística dd/mm vs mm/dd
        if a > 12 and 1 <= b <= 12:
            try: return datetime.date(c, b, a)
            except: return None
        if b > 12 and 1 <= a <= 12:
            try: return datetime.date(c, a, b)
            except: return None
        for fmt in ('%d/%m/%Y', '%m/%d/%Y'):
            try: return datetime.datetime.strptime(f"{a:02d}/{b:02d}/{c:04d}", fmt).date()
            except: continue
        return None

    # Fallback con dateutil
    for dayfirst in (True, False):
        try:
            dt = parse(s, dayfirst=dayfirst, yearfirst=False)
            return dt.date()
        except: continue
    return None

def format_date_ddmmyyyy(dt):
    """Devuelve 'dd/mm/yyyy' o 'Sin fecha'."""
    if not isinstance(dt, (datetime.date, datetime.datetime)) or dt is None:
        return 'Sin fecha'
    if isinstance(dt, datetime.datetime):
        dt = dt.date()
    return dt.strftime('%d/%m/%Y')

# =========================================================
# LIMPIEZA DE CORREOS (Desde la columna 'Correos')
# =========================================================
EMAIL_REGEX = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')

def extract_emails(correo_info: str) -> list[str]:
    """
    Devuelve lista de correos válidos a partir del contenido de la celda 'Correos'.
    - Soporta 'mailto:', múltiples correos separados por ; , / espacios, y texto mezclado.
    - Ignora texto adicional y números pegados a los correos.
    """
    if not correo_info:
        return []
    s = str(correo_info)
    s = s.replace('mailto:', ' ')
    emails = EMAIL_REGEX.findall(s)
    seen = set()
    result = []
    for e in emails:
        e2 = e.strip()
        if e2 and e2.lower() not in seen:
            seen.add(e2.lower())
            result.append(e2)
    return result

# =========================================================
# DIBUJO EN PLANTILLAS (SOAT / RTM)
# =========================================================
def _try_load_font(size=100):
    """Carga una fuente negrita (arialbd) con fallback a arial o default."""
    candidates = [
        r"C:\RPA\CE0864_Consulta_Runt_Simit\Aut. Runt x Cedula y Placa\Fuente\arialbd.ttf",
        "arialbd.ttf",
        "arial.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()

def add_text_to_image_soat(image_path, placa, fecha_str, output_path):
    """
    SOAT: texto negro con borde blanco.
    Coordenadas fijas (x=490, y=1750) pensadas para tu plantilla SOAT (4252x4661).
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"No existe la plantilla: {image_path}")

    img = Image.open(image_path).convert("RGB")
    W, H = img.size
    draw = ImageDraw.Draw(img)

    # Tamaño grande para SOAT (plantilla alta)
    font = _try_load_font(100)

    text = (
        "El Seguro Obligatorio de Accidentes de Tránsito (SOAT) del \n"
        f"Vehículo: {placa} se encuentra en Estado: NO VIGENTE\n"
        f"Fecha de vigencia: {fecha_str}"
    )

    # Coordenadas fijas (ajústalas si cambias plantilla)
    x, y = 490, 1750

    draw.multiline_text(
        (x, y),
        text,
        fill=(0, 0, 0),               # negro
        font=font,
        align="left",
        spacing=3,
        stroke_width=2,
        stroke_fill=(255, 255, 255)   # borde blanco
    )

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    img.save(output_path, format="JPEG")
    return output_path

def add_text_to_image_rtm(image_path, placa, fecha_str, output_path):
    """
    RTM (fondo blanco): texto negro con borde blanco en (520, 1800).
    Si la coordenada Y queda fuera del alto de la imagen, ajusta a 70% del alto.
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"No existe la plantilla: {image_path}")

    img = Image.open(image_path).convert("RGB")
    W, H = img.size
    draw = ImageDraw.Draw(img)

    font = _try_load_font(13)

    text = (
        "La Revisión Técnico Mecánica (RTM) del \n"
        f"Vehículo: {placa} Se encuentra en Estado: NO VIGENTE\n"
        f"Fecha de vigencia: {fecha_str}"
    )

    x, y = 88, 309
    

    draw.multiline_text(
        (x, y),
        text,
        fill=(0, 0, 0),               # negro
        font=font,
        align="left",
        spacing=2,
        stroke_width=2,
        stroke_fill=(255, 255, 255)   # borde blanco
    )

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    img.save(output_path, format="JPEG")
    return output_path


# =========================================================
# OUTLOOK: envío a lista de destinatarios (TO/CC/BCC)
# =========================================================
def send_email_to_recipients(recipients: list[str], subject: str, body_html: str,
                             attachments: list[str] | None = None,
                             cc: list[str] | None = None,
                             bcc: list[str] | None = None):
    """
    Envía correo vía Outlook a 'recipients' (TO), opcionalmente con CC/BCC y adjuntos.
    """
    recipients = [r for r in (recipients or []) if r]
    if not recipients:
        raise ValueError("No hay destinatarios válidos para enviar el correo (TO).")

    try:
        outlook = win32.Dispatch('outlook.application')
    except Exception as e:
        raise RuntimeError("No se pudo iniciar Outlook. Verifica que Outlook esté instalado/abierto.") from e

    mail = outlook.CreateItem(0)

    # Destinatarios
    mail.To  = "; ".join(recipients)
    mail.CC  = "; ".join(cc or [])
    mail.BCC = "; ".join(bcc or [])

    mail.Subject  = subject
    mail.HTMLBody = body_html

    # Adjuntos existentes
    for att in (attachments or []):
        if att and os.path.exists(att):
            mail.Attachments.Add(att)

    # Log simple
    try:
        log_path = os.path.join(BASE_IMG_DIR, "envios_log.txt")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[ENVIO] TO: {mail.To} | CC: {mail.CC} | BCC: {mail.BCC} | Asunto: {subject}\n")
    except Exception:
        pass

    mail.Send()

# =========================================================
# CONSTRUCCIÓN DE NOTIFICACIONES POR MÉTRICA (texto + fecha)
# =========================================================
def build_notifications_by_metrics():
    """
    Devuelve lista [(tipo, fecha_str, placa, correo_info)] SOLO cuando se cumplen AMBAS condiciones:
      - SOAT: Estado del SOAT ∈ {'NO VIGENTE', ...}  Y  Vigencia SOAT < hoy
      - RTM : Vigencia RTM   ∈ {'NO', 'NO VIGENTE', ...}  Y  Fecha vigencia RTM < hoy
    Se EXCLUYEN explícitamente 'SIN INFO' y 'ERROR'.
    Las fechas se formatean para mostrar en correo/imagen.
    """
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=EXCEL_SHEET, engine='openpyxl')
    except Exception as e:
        messagebox.showerror("Error al leer Excel", f"No se pudo leer:\n{EXCEL_PATH}\n\nDetalle: {e}")
        return []

    notifications = []
    today = datetime.date.today()

    # Estados válidos post-normalización
    SOAT_INVALID_STATES = {'NO VIGENTE', 'NO VIGENTE.'}
    RTM_INVALID_STATES  = {'NO', 'NO.', 'NO VIGENTE', 'NO VIGENTE.'}
    EXCLUDE_STATES      = {'SIN INFO', 'SIN INFO.', 'ERROR'}

    for _, row in df.iterrows():
        placa   = str(row.get(PLACA_COLUMN, "")).strip()
        correo  = str(row.get(CORREOS_COL, "")).strip()

        # SOAT
        soat_raw        = row.get(ESTADO_SOAT_COL, "")
        soat_estado     = norm_text_exact(soat_raw)
        soat_fecha_obj  = parse_mixed_date(row.get(FECHA_SOAT_COL, None))
        soat_excluir    = (soat_estado in EXCLUDE_STATES)

        # RTM
        rtm_raw         = row.get(VIGENCIA_RTM_COL, "")
        rtm_estado      = norm_text_exact(rtm_raw)
        rtm_fecha_obj   = parse_mixed_date(row.get(FECHA_RTM_COL, None))
        rtm_excluir     = (rtm_estado in EXCLUDE_STATES)

        # Métrica SOAT
        soat_send = (
            not soat_excluir and
            (soat_estado in SOAT_INVALID_STATES) and
            isinstance(soat_fecha_obj, datetime.date) and
            soat_fecha_obj < today
        )

        # Métrica RTM (acepta 'NO' y 'NO VIGENTE')
        rtm_send = (
            not rtm_excluir and
            (rtm_estado in RTM_INVALID_STATES) and
            isinstance(rtm_fecha_obj, datetime.date) and
            rtm_fecha_obj < today
        )

        if soat_send:
            notifications.append(("SOAT", format_date_ddmmyyyy(soat_fecha_obj), placa, correo))
        if rtm_send:
            notifications.append(("RTM", format_date_ddmmyyyy(rtm_fecha_obj), placa, correo))

    return notifications

# =========================================================
# INTERFAZ: REFRESCAR (solo inválidos) y NOTIFICAR (solo envía)
# =========================================================
def load_alerts_only_invalids():
    """Devuelve la lista de inválidos para mostrar (usa la misma métrica que se notifica)."""
    return build_notifications_by_metrics()

def populate_tree(alerts):
    for item in tree.get_children():
        tree.delete(item)
    for alert in alerts:
        tree.insert('', 'end', values=alert)  # (Tipo, Fecha, Placa, Correo)

def refresh():
    """
    Refresca la tabla mostrando SOLO los inválidos según métrica:
      - SOAT: Estado == 'NO VIGENTE' y fecha SOAT < hoy
      - RTM : Vigencia ∈ {'NO','NO VIGENTE'} y fecha RTM < hoy
    """
    invalids = load_alerts_only_invalids()
    populate_tree(invalids)
    status_label.config(text=f"Mostrando inválidos: {len(invalids)} registro(s)")
    progress['value'] = 0

def notify():
    """
    Envía UNO POR UNO solo cuando se cumplen las métricas.
    Ahora toma los correos desde la columna 'Correos' de cada fila.
    >>> NO repuebla la tabla (solo envía).
    """
    alerts = build_notifications_by_metrics()

    if not alerts:
        messagebox.showinfo("Sin coincidencias", "No hay registros que cumplan las métricas (texto + fecha).")
        return

    total = len(alerts)
    enviados = 0
    errores = 0
    ultimo_error = ""

    progress['maximum'] = total
    progress['value'] = 0

    for idx, (tipo, fecha_str, placa, correo_info) in enumerate(alerts, start=1):
        status_label.config(
            text=f"Enviando {idx}/{total} | Enviados: {enviados} | Errores: {errores} | Placa: {placa} | Tipo: {tipo}"
        )
        progress['value'] = idx - 1
        root.update_idletasks()

        try:
            # Extrae los correos válidos de la celda 'Correos'
            recipients = extract_emails(correo_info)
            if not recipients:
                errores += 1
                ultimo_error = f"Fila sin correos válidos (Placa {placa}). Contenido: {correo_info}"
                continue

            if tipo.upper() == "SOAT":
                add_text_to_image_soat(PLANTILLA_SOAT, placa, fecha_str, SALIDA_SOAT)
                html = f"""
                <html>
                  <body>
                    <p><strong>Alerta: SOAT NO VIGENTE</strong></p>
                    <p>Vehículo: <strong>{placa}</strong></p>
                    <p>Fecha de vigencia: <strong>{fecha_str}</strong></p>
                    <p>Se adjunta imagen.</p>
                  </body>
                </html>
                """
                send_email_to_recipients(
                    recipients=recipients,
                    subject=f"⚠ SOAT NO VIGENTE | Placa {placa} | {fecha_str}",
                    body_html=html,
                    attachments=[SALIDA_SOAT],
                    bcc=DEFAULT_BCC  # opcional
                )

            elif tipo.upper() == "RTM":
                add_text_to_image_rtm(PLANTILLA_RTM, placa, fecha_str, SALIDA_RTM)
                html = f"""
                <html>
                  <body>
                    <p><strong>Alerta: RTM NO VÁLIDO</strong></p>
                    <p>Vehículo: <strong>{placa}</strong></p>
                    <p>Fecha de vigencia: <strong>{fecha_str}</strong></p>
                    <p>Se adjunta imagen.</p>
                  </body>
                </html>
                """
                send_email_to_recipients(
                    recipients=recipients,
                    subject=f"⚠ RTM NO VÁLIDO | Placa {placa} | {fecha_str}",
                    body_html=html,
                    attachments=[SALIDA_RTM],
                    bcc=DEFAULT_BCC  # opcional
                )

            enviados += 1
            progress['value'] = idx
            status_label.config(
                text=f"Enviando {idx}/{total} | Enviados: {enviados} | Errores: {errores} | Placa: {placa} | Tipo: {tipo}"
            )
            root.update_idletasks()
            time.sleep(0.2)

        except Exception as e:
            errores += 1
            ultimo_error = str(e)
            status_label.config(
                text=f"Enviando {idx}/{total} | Enviados: {enviados} | Errores: {errores} | Error: {ultimo_error}"
            )
            root.update_idletasks()
            time.sleep(0.2)

    status_label.config(text=f"Finalizado | Total: {total} | Enviados: {enviados} | Errores: {errores}")
    progress['value'] = total
    root.update_idletasks()

    if errores:
        messagebox.showwarning(
            "Envío con errores",
            f"Procesados: {total}\nEnviados: {enviados}\nErrores: {errores}\n\nÚltimo error: {ultimo_error}"
        )
    else:
        messagebox.showinfo("Éxito", f"Procesados: {total}\nEnviados: {enviados}\nErrores: {errores}")

# =========================================================
# INTERFAZ
# =========================================================
root = tk.Tk()
root.title("Alerta de fechas de vigencia")
root.resizable(False, False)

tree = ttk.Treeview(root, columns=('Tipo', 'Fecha', 'Placa', 'Correo'), show='headings')
tree.heading('Tipo', text='Tipo')
tree.heading('Fecha', text='Fecha registrada')
tree.heading('Placa', text='Placa')
tree.heading('Correo', text='Correo')

tree.column('Tipo', width=140, anchor='center')
tree.column('Fecha', width=180, anchor='center')
tree.column('Placa', width=140, anchor='center')
tree.column('Correo', width=320, anchor='w')

tree.pack(padx=20, pady=10, fill='both')

btn_frame = tk.Frame(root)
btn_frame.pack(fill='x', padx=20, pady=10)

# Botón Refrescar: ahora lista SOLO inválidos
refresh_btn = tk.Button(
    btn_frame,
    text="Refrescar",
    command=refresh,
    font=("Arial", 11, 'bold'),
    bg="#00A36C",
    fg="white",
    activebackground="#00C387",
    activeforeground="black",
    bd=4,
    width=24
)
refresh_btn.pack(side='left')

# Ícono (opcional)
try:
    icono = PhotoImage(file=r"C:\RPA\CE0864_Consulta_Runt_Simit\Aut. Runt x Cedula y Placa\img\iconOutlook.png")
except Exception:
    icono = None

# Botón Notificar: solo envía
notify_btn = tk.Button(
    btn_frame,
    text="Notificar inválidos (enviar)",
    command=notify,
    font=("Arial", 12, 'bold'),
    bg="#4169FF",
    fg="white",
    activebackground="lightblue",
    activeforeground="black",
    bd=5,
    height=30,
    width=260,
    image=icono if icono else None,
    compound="left"
)
notify_btn.pack(side='right')

status_frame = tk.Frame(root)
status_frame.pack(fill='x', padx=20, pady=(0, 12))

status_label = tk.Label(status_frame, text="Listo", anchor='w')
status_label.pack(side='left')

progress = ttk.Progressbar(status_frame, orient='horizontal', mode='determinate', length=240)
progress.pack(side='right')

# Inicializa la tabla mostrando SOLO inválidos
refresh()

root.mainloop()
