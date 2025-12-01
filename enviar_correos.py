#BACKUP FILE BEFORE RUNNING!
import pandas as pd
import win32com.client as win32
from pathlib import Path

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

RUTA = Path(r"C:\Users\leona\OneDrive\Documentos\RASPBERRY")

ARCHIVO_BAJAS = RUTA / "2025-11-28_consolidado_de_bajas.xlsx"
ARCHIVO_ALTAS = RUTA / "2025-11-28_Template_de_Altas_generado.xlsx"

# ⚠️ IMPORTANTE:
# True  -> SOLO crea borradores
# False -> ENVÍA automáticamente
MODO_BORRADOR = False

ENVIAR_CORREOS_BAJAS = True
ENVIAR_CORREOS_ALTAS = True

NOMBRE_REMITENTE = "Raspberry Servicios"

# Correo que recibirá la confirmación de resumen
CORREO_CONFIRMACION = "leonardojaviercarranza@gmail.com"
# ============================================================
# LIMPIAR BANDEJA DE SALIDA (OPCIONAL)
# ============================================================

# Si pones esto en True, ese run del script SOLO limpiará Bandeja de salida
LIMPIAR_BANDEJA_SALIDA = False  # ⚠️ CAMBIA A True SOLO CUANDO QUIERAS LIMPIAR

# "borrar" → elimina los correos en cola
# "enviar" → intenta mandarlos todos
MODO_LIMPIEZA = "borrar"   # "borrar" o "enviar"


# ============================================================
# OUTLOOK
# ============================================================

def obtener_outlook():
    return win32.Dispatch("Outlook.Application")


def limpiar_bandeja_salida(outlook_app, modo="borrar"):
    """
    modo = "borrar"  -> borra todos los correos de la Bandeja de salida
    modo = "enviar"  -> intenta enviar todos los correos de la Bandeja de salida
    """
    mapi = outlook_app.GetNamespace("MAPI")
    # 4 = olFolderOutbox
    outbox = mapi.GetDefaultFolder(4)

    total = outbox.Items.Count
    print(f"📂 Correos en Bandeja de salida: {total}")

    if total == 0:
        print("No hay correos en la Bandeja de salida.")
        return

    for item in list(outbox.Items):
        try:
            if modo == "enviar":
                subject = getattr(item, "Subject", "(sin asunto)")
                item.Send()
                print(f"✉ Enviado desde bandeja de salida: {subject}")
            elif modo == "borrar":
                subject = getattr(item, "Subject", "(sin asunto)")
                item.Delete()
                print(f"🗑 Borrado desde bandeja de salida: {subject}")
            else:
                raise ValueError("modo debe ser 'borrar' o 'enviar'")
        except Exception as e:
            print(f"⚠ Error con un ítem de la Bandeja de salida: {e}")

    print("✅ Limpieza de Bandeja de salida terminada.\n")


# ============================================================
# COLUMNAS NECESARIAS
# ============================================================

# BAJAS (solo correo + póliza)
COL_CORREO_BAJAS      = "Correo"
COL_POLIZA_BAJAS      = "Nº ref.externo"

# ALTAS (solo correo + póliza)
COL_CORREO_ALTAS      = "Correo"
COL_POLIZA_ALTAS      = "Nº referencia externo"


# ============================================================
# CORREOS DE BAJAS
# ============================================================
def enviar_correos_bajas(outlook_app):
    print("\n========== ENVIANDO CORREOS DE BAJAS ==========\n")

    df_bajas_raw = pd.read_excel(ARCHIVO_BAJAS)
    df_bajas_raw.columns = df_bajas_raw.columns.str.strip()

    for col in [COL_CORREO_BAJAS, COL_POLIZA_BAJAS]:
        if col not in df_bajas_raw.columns:
            raise ValueError(f"❌ Falta columna '{col}' en {ARCHIVO_BAJAS.name}")

    # Total de registros con póliza (aunque no tengan correo)
    total_registros_bajas = len(df_bajas_raw[df_bajas_raw[COL_POLIZA_BAJAS].notna()])

    # Filtrar solo filas con correo y póliza válidos (estos sí se envían)
    df_bajas = df_bajas_raw[
        df_bajas_raw[COL_CORREO_BAJAS].notna()
        & df_bajas_raw[COL_CORREO_BAJAS].astype(str).str.strip().ne("")
        & df_bajas_raw[COL_POLIZA_BAJAS].notna()
    ]

    total_bajas_enviadas = len(df_bajas)
    print(f"Total de bajas con correo: {total_bajas_enviadas} / {total_registros_bajas}")

    # Si no hay bajas con correo → enviar correo de resumen con 0
    if df_bajas.empty:
        print("No hay bajas con correo. Se enviará correo de confirmación con 0 registros.")

        subject_conf = "Resumen de envío de BAJAS – 0 correos enviados"
        body_conf = f"""
Hola,

Se ejecutó el proceso de BAJAS, pero no se encontró ningún registro con correo válido.

• Archivo procesado: {ARCHIVO_BAJAS.name}
• Total de correos de BAJA enviados: 0/{total_registros_bajas} colaboradores

Modo de ejecución: {"BORRADOR (solo se hubieran abierto en Outlook)" if MODO_BORRADOR else "ENVÍO REAL (habrían sido enviados por Outlook)"}

Saludos,
{NOMBRE_REMITENTE}
"""

        mail_conf = outlook_app.CreateItem(0)
        mail_conf.To = CORREO_CONFIRMACION
        mail_conf.Subject = subject_conf
        mail_conf.Body = body_conf

        if MODO_BORRADOR:
            mail_conf.Display()
            print(f"✔ Borrador correo de CONFIRMACIÓN BAJAS → {CORREO_CONFIRMACION}")
        else:
            mail_conf.Send()
            print(f"✉ Correo de CONFIRMACIÓN BAJAS enviado → {CORREO_CONFIRMACION}")

        print("\n✔ Finalizado proceso de BAJAS (sin registros con correo).\n")
        return

    # Si hay registros → mandar correo individual por persona
    for _, row in df_bajas.iterrows():
        correo = str(row[COL_CORREO_BAJAS]).strip()
        poliza = str(row[COL_POLIZA_BAJAS]).strip()

        subject = f"Notificación de baja de póliza {poliza}"
        body = f"""
Hola,

Se ha registrado una BAJA de tu seguro de automóvil.

• Número de póliza: {poliza}
 
Texto...

Saludos,
{NOMBRE_REMITENTE}
"""

        mail = outlook_app.CreateItem(0)
        mail.To = correo
        mail.Subject = subject
        mail.Body = body

        if MODO_BORRADOR:
            mail.Display()
            print(f"✔ Borrador BAJA → {correo} | póliza {poliza}")
        else:
            mail.Send()
            print(f"✉ BAJA enviada → {correo} | póliza {poliza}")

    # Al terminar, enviar correo de confirmación de BAJAS
    subject_conf = "Confirmación de envío de BAJAS"
    body_conf = f"""
Hola,

Se ha ejecutado el proceso de envío de BAJAS de seguros de automóvil.

• Archivo procesado: {ARCHIVO_BAJAS.name}
• Total de correos de BAJA enviados: {total_bajas_enviadas}/{total_registros_bajas} colaboradores

Texto...

Saludos,
{NOMBRE_REMITENTE}
"""

    mail_conf = outlook_app.CreateItem(0)
    mail_conf.To = CORREO_CONFIRMACION
    mail_conf.Subject = subject_conf
    mail_conf.Body = body_conf

    if MODO_BORRADOR:
        mail_conf.Display()
        print(f"✔ Borrador correo de CONFIRMACIÓN BAJAS → {CORREO_CONFIRMACION}")
    else:
        mail_conf.Send()
        print(f"✉ Correo de CONFIRMACIÓN BAJAS enviado → {CORREO_CONFIRMACION}")

    print("\n✔ Finalizado proceso de BAJAS.\n")


# ============================================================
# CORREOS DE ALTAS
# ============================================================
def enviar_correos_altas(outlook_app):
    print("\n========== ENVIANDO CORREOS DE ALTAS ==========\n")

    df_altas_raw = pd.read_excel(ARCHIVO_ALTAS)
    df_altas_raw.columns = df_altas_raw.columns.str.strip()

    for col in [COL_CORREO_ALTAS, COL_POLIZA_ALTAS]:
        if col not in df_altas_raw.columns:
            raise ValueError(f"❌ Falta columna '{col}' en {ARCHIVO_ALTAS.name}")

    # Total de registros con póliza (aunque no tengan correo)
    total_registros_altas = len(df_altas_raw[df_altas_raw[COL_POLIZA_ALTAS].notna()])

    # Filtrar solo filas con correo y póliza válidos (estos sí se envían)
    df_altas = df_altas_raw[
        df_altas_raw[COL_CORREO_ALTAS].notna()
        & df_altas_raw[COL_CORREO_ALTAS].astype(str).str.strip().ne("")
        & df_altas_raw[COL_POLIZA_ALTAS].notna()
    ]

    total_altas_enviadas = len(df_altas)
    print(f"Total de altas con correo: {total_altas_enviadas} / {total_registros_altas}")

    # Si no hay altas con correo → enviar correo de resumen con 0
    if df_altas.empty:
        print("No hay altas con correo. Se enviará correo de confirmación con 0 registros.")

        subject_conf = "Resumen de envío de ALTAS – 0 correos enviados"
        body_conf = f"""
Hola,

• Archivo procesado: {ARCHIVO_ALTAS.name}
• Total de correos de ALTA enviados: 0/{total_registros_altas} colaboradores

Se ejecutó el proceso de ALTAS, pero no se encontró ningún registro con correo válido.

Texto...

Saludos,
{NOMBRE_REMITENTE}
"""

        mail_conf = outlook_app.CreateItem(0)
        mail_conf.To = CORREO_CONFIRMACION
        mail_conf.Subject = subject_conf
        mail_conf.Body = body_conf

        if MODO_BORRADOR:
            mail_conf.Display()
            print(f"✔ Borrador correo de CONFIRMACIÓN ALTAS → {CORREO_CONFIRMACION}")
        else:
            mail_conf.Send()
            print(f"✉ Correo de CONFIRMACIÓN ALTAS enviado → {CORREO_CONFIRMACION}")

        print("\n✔ Finalizado proceso de ALTAS (sin registros con correo).\n")
        return

    # Si hay registros → mandar correo individual por persona
    for _, row in df_altas.iterrows():
        correo = str(row[COL_CORREO_ALTAS]).strip()
        poliza = str(row[COL_POLIZA_ALTAS]).strip()

        subject = f"Notificación de alta de póliza {poliza}"
        body = f"""
Hola,

Se ha registrado el ALTA de tu seguro de automóvil.

• Número de póliza: {poliza}

Texto...

Saludos,
{NOMBRE_REMITENTE}
"""

        mail = outlook_app.CreateItem(0)
        mail.To = correo
        mail.Subject = subject
        mail.Body = body

        if MODO_BORRADOR:
            mail.Display()
            print(f"✔ Borrador ALTA → {correo} | póliza {poliza}")
        else:
            mail.Send()
            print(f"✉ ALTA enviada → {correo} | póliza {poliza}")

    # Al terminar, enviar correo de confirmación de ALTAS
    subject_conf = "Confirmación de envío de ALTAS"
    body_conf = f"""
Hola,

Se ha ejecutado el proceso de envío de ALTAS de seguros de automóvil.

• Archivo procesado: {ARCHIVO_ALTAS.name}
• Total de correos de ALTA enviados: {total_altas_enviadas}/{total_registros_altas} colaboradores

Texto...

Saludos,
{NOMBRE_REMITENTE}
"""

    mail_conf = outlook_app.CreateItem(0)
    mail_conf.To = CORREO_CONFIRMACION
    mail_conf.Subject = subject_conf
    mail_conf.Body = body_conf

    if MODO_BORRADOR:
        mail_conf.Display()
        print(f"✔ Borrador correo de CONFIRMACIÓN ALTAS → {CORREO_CONFIRMACION}")
    else:
        mail_conf.Send()
        print(f"✉ Correo de CONFIRMACIÓN ALTAS enviado → {CORREO_CONFIRMACION}")

    print("\n✔ Finalizado proceso de ALTAS.\n")

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    outlook = obtener_outlook()

    # 🧹 MODO LIMPIEZA DE BANDEJA DE SALIDA
    if LIMPIAR_BANDEJA_SALIDA:
        limpiar_bandeja_salida(outlook, modo=MODO_LIMPIEZA)
    else:
        # Modo normal: envío de correos de bajas/altas
        if ENVIAR_CORREOS_BAJAS:
            enviar_correos_bajas(outlook)

        if ENVIAR_CORREOS_ALTAS:
            enviar_correos_altas(outlook)

        print("🎉 Script de correos Raspberry finalizado.")
