import streamlit as st
import pandas as pd
from pandas import DateOffset
from io import BytesIO
from datetime import datetime

# =========================================
# FUNCIONES DE PROCESO (BAJAS Y ALTAS)
# =========================================

def procesar_bajas(parque, cancelacion, cancelado, nomina):
    # -------------------- CREAR COLUMNAS NUEVAS EN NOMINA --------------------
    columnas_nuevas = [
        "Cancelado",
        "Fecha de cancelado",
        "Cancelación",
        "Fecha de Cancelación",
        "Anulados GNP"
    ]

    for col in columnas_nuevas:
        if col not in nomina.columns:
            nomina[col] = None

    # -------------------- PASO 2: LLENAR "CANCELADO" Y "FECHA DE CANCELADO" --------------------
    COL_LLAVE_NOMINA    = "Nº ref.externo"
    COL_LLAVE_CANCELADO = "No. Póliza"
    COL_FECHA_BAJA_OP   = "Fecha de Baja Operativa"

    # 1) Normalizar llaves como texto
    for df, col in [(nomina, COL_LLAVE_NOMINA), (cancelado, COL_LLAVE_CANCELADO)]:
        df[col] = df[col].astype(str).str.strip()

    # 2) Preparar cancelado: convertir fechas y eliminar duplicados
    cancelado_tmp = cancelado.copy()
    cancelado_tmp[COL_FECHA_BAJA_OP] = pd.to_datetime(
        cancelado_tmp[COL_FECHA_BAJA_OP], errors="coerce"
    )

    cancelado_tmp = (
        cancelado_tmp
        .sort_values(COL_FECHA_BAJA_OP)
        .drop_duplicates(subset=COL_LLAVE_CANCELADO, keep="first")
        .set_index(COL_LLAVE_CANCELADO)
    )

    # Serie tipo BUSCARV: póliza → póliza
    serie_poliza = pd.Series(cancelado_tmp.index, index=cancelado_tmp.index)

    # 3) Llenar "Cancelado"
    nomina["Cancelado"] = nomina[COL_LLAVE_NOMINA].map(serie_poliza)

    # 4) Llenar "Fecha de cancelado"
    nomina["Fecha de cancelado"] = nomina[COL_LLAVE_NOMINA].map(
        cancelado_tmp[COL_FECHA_BAJA_OP]
    )

    # -------------------- BACKUP ANTES DE ELIMINAR FILAS --------------------
    df_nomina_cancelaciones_antes_de_eliminar = nomina.copy()
    df_nomina_cancelaciones_antes_de_eliminar = df_nomina_cancelaciones_antes_de_eliminar.drop(
        columns=["Cancelado", "Fecha de cancelado"],
        errors="ignore"
    )

    # -------------------- FILTRAR SOLO FILAS DONDE SI EXISTE CANCELADO --------------------
    nomina["Cancelado"] = nomina["Cancelado"].replace("", pd.NA)
    nomina = nomina[nomina["Cancelado"].notna()]

    # -------------------- ELIMINAR FILAS CON VALORES NEGATIVOS O 0 --------------------
    columnas_valores = ["SaldoIni", "DesctPer", "SaldoFin"]

    for col in columnas_valores:
        if col in nomina.columns:
            nomina[col] = pd.to_numeric(nomina[col], errors="coerce")

    nomina = nomina[
        (nomina["SaldoIni"] > 0) &
        (nomina["DesctPer"] > 0) &
        (nomina["SaldoFin"] > 0)
    ]

    # -------------------- FORMATEAR TODAS LAS FECHAS --------------------
    columnas_fecha = [
        "FePago",
        "InicioPer",
        "FinPer",
        "InicioVal",
        "FinVal",
        "FinPrevistoPréstamo",
        "Fecha de cancelado",
    ]

    for col in columnas_fecha:
        if col in nomina.columns:
            nomina[col] = pd.to_datetime(
                nomina[col], errors="coerce"
            ).dt.strftime("%d/%m/%Y")

    # -------------------- ELIMINAR COLUMNAS TOTALMENTE VACÍAS --------------------
    columnas_vacias = [col for col in nomina.columns if nomina[col].isna().all()]
    if columnas_vacias:
        nomina.drop(columns=columnas_vacias, inplace=True)

    # ============================================================
    # PASO 3: LLENAR "Cancelación" Y "Fecha de Cancelación"
    # ============================================================
    df_cancelacion_nomina = df_nomina_cancelaciones_antes_de_eliminar.copy()

    COL_LLAVE_NOMINA_CANC  = "Nº ref.externo"
    COL_LLAVE_CANCELACION  = "No. Póliza"
    COL_FECHA_BAJA_OP_CANC = "Fecha de Baja Operativa"

    for df_tmp, col in [
        (df_cancelacion_nomina, COL_LLAVE_NOMINA_CANC),
        (cancelacion,          COL_LLAVE_CANCELACION)
    ]:
        df_tmp[col] = df_tmp[col].astype(str).str.strip()

    cancelacion_tmp = cancelacion.copy()
    cancelacion_tmp[COL_FECHA_BAJA_OP_CANC] = pd.to_datetime(
        cancelacion_tmp[COL_FECHA_BAJA_OP_CANC], errors="coerce"
    )

    cancelacion_tmp = (
        cancelacion_tmp
        .sort_values(COL_FECHA_BAJA_OP_CANC)
        .drop_duplicates(subset=COL_LLAVE_CANCELACION, keep="first")
        .set_index(COL_LLAVE_CANCELACION)
    )

    serie_poliza_canc = pd.Series(cancelacion_tmp.index, index=cancelacion_tmp.index)

    df_cancelacion_nomina["Cancelación"] = df_cancelacion_nomina[COL_LLAVE_NOMINA_CANC].map(
        serie_poliza_canc
    )
    df_cancelacion_nomina["Fecha de Cancelación"] = df_cancelacion_nomina[COL_LLAVE_NOMINA_CANC].map(
        cancelacion_tmp[COL_FECHA_BAJA_OP_CANC]
    )

    df_cancelacion_nomina["Cancelación"] = df_cancelacion_nomina["Cancelación"].replace("", pd.NA)
    df_cancelacion_nomina = df_cancelacion_nomina[df_cancelacion_nomina["Cancelación"].notna()]

    # Limpiar negativos/ceros en Cancelación
    columnas_valores_cancelacion = ["SaldoIni", "DesctPer", "SaldoFin"]
    for col in columnas_valores_cancelacion:
        if col in df_cancelacion_nomina.columns:
            df_cancelacion_nomina[col] = pd.to_numeric(
                df_cancelacion_nomina[col], errors="coerce"
            )

    df_cancelacion_nomina = df_cancelacion_nomina[
        (df_cancelacion_nomina["SaldoIni"]  > 0) & 
        (df_cancelacion_nomina["DesctPer"] > 0) & 
        (df_cancelacion_nomina["SaldoFin"] > 0)
    ]

    # Formatear fechas Cancelación
    columnas_fechas_cancelacion = [
        "FePago",
        "InicioPer",
        "FinPer",
        "InicioVal",
        "FinVal",
        "FinPrevistoPréstamo",
        "Fecha de Cancelación"
    ]

    for col in columnas_fechas_cancelacion:
        if col in df_cancelacion_nomina.columns:
            df_cancelacion_nomina[col] = pd.to_datetime(
                df_cancelacion_nomina[col], errors="coerce"
            ).dt.strftime("%d/%m/%Y")

    columnas_vacias_canc = [col for col in df_cancelacion_nomina.columns
                            if df_cancelacion_nomina[col].isna().all()]
    if columnas_vacias_canc:
        df_cancelacion_nomina.drop(columns=columnas_vacias_canc, inplace=True)

    # ============================================================
    # PASO 4: LLENAR "Anulados GNP" USANDO PARQUE (HOJA ANULADAS)
    # ============================================================
    df_anulados_nomina = df_nomina_cancelaciones_antes_de_eliminar.copy()

    COL_LLAVE_NOMINA_ANU  = "Nº ref.externo"
    COL_LLAVE_PARQUE_ANU  = "CDNUMPOL"

    for df_tmp, col in [
        (df_anulados_nomina, COL_LLAVE_NOMINA_ANU),
        (parque,             COL_LLAVE_PARQUE_ANU)
    ]:
        df_tmp[col] = df_tmp[col].astype(str).str.strip()

    parque_tmp = parque.copy()
    parque_tmp = (
        parque_tmp
        .drop_duplicates(subset=COL_LLAVE_PARQUE_ANU, keep="first")
        .set_index(COL_LLAVE_PARQUE_ANU)
    )

    serie_anulados = pd.Series(parque_tmp.index, index=parque_tmp.index)

    df_anulados_nomina["Anulados GNP"] = df_anulados_nomina[COL_LLAVE_NOMINA_ANU].map(
        serie_anulados
    )

    df_anulados_nomina["Anulados GNP"] = df_anulados_nomina["Anulados GNP"].replace("", pd.NA)
    df_anulados_nomina = df_anulados_nomina[df_anulados_nomina["Anulados GNP"].notna()]

    columnas_valores_anulados = ["SaldoIni", "DesctPer", "SaldoFin"]
    for col in columnas_valores_anulados:
        if col in df_anulados_nomina.columns:
            df_anulados_nomina[col] = pd.to_numeric(df_anulados_nomina[col], errors="coerce")

    if all(col in df_anulados_nomina.columns for col in columnas_valores_anulados):
        df_anulados_nomina = df_anulados_nomina[
            (df_anulados_nomina["SaldoIni"]  > 0) & 
            (df_anulados_nomina["DesctPer"] > 0) & 
            (df_anulados_nomina["SaldoFin"] > 0)
        ]

    # Borrar columnas de Cancelación en Anulados
    df_anulados_nomina = df_anulados_nomina.drop(
        columns=["Cancelación", "Fecha de Cancelación"],
        errors="ignore"
    )

    # ============================================================
    # PASO 5: CONSOLIDADO DE BAJAS
    # ============================================================
    df_cancelado_consol   = nomina.copy()
    df_cancelacion_consol = df_cancelacion_nomina.copy()
    df_anulados_consol    = df_anulados_nomina.copy()

    columnas_originales = df_nomina_cancelaciones_antes_de_eliminar.columns.tolist()

    df_cancelado_consol   = df_cancelado_consol.reindex(columns=columnas_originales)
    df_cancelacion_consol = df_cancelacion_consol.reindex(columns=columnas_originales)
    df_anulados_consol    = df_anulados_consol.reindex(columns=columnas_originales)

    df_cancelado_consol["Tipo_baja"]   = "Cancelado"
    df_cancelacion_consol["Tipo_baja"] = "Cancelación"
    df_anulados_consol["Tipo_baja"]    = "Anulado GNP"

    consolidado_bajas = pd.concat(
        [df_cancelado_consol, df_cancelacion_consol, df_anulados_consol],
        ignore_index=True
    )

    cols_a_eliminar = [
        "Cancelado",
        "Fecha de cancelado",
        "Cancelación",
        "Fecha de Cancelación",
        "Anulados GNP"
    ]
    consolidado_bajas = consolidado_bajas.drop(columns=cols_a_eliminar, errors="ignore")

    return consolidado_bajas


def procesar_altas(parque_vigentes, activos, nomina, template_altas):
    # ==============================
    # PASO 1 — LIMPIAR ACTIVOS
    # ==============================
    col_institucion = None
    for c in activos.columns:
        if c.strip().lower() == "institución".lower():
            col_institucion = c
            break
    if col_institucion is None:
        raise ValueError("No encontré la columna de Institución en Activos.xlsx")

    instituciones_a_borrar = [
        "RASPBERRY SERVICES",
        "AKKY ONLINE SOLUTIONS",
        "NETWORK INFORMATION CENTER",
        "URBANIZADORA PARA EL DESARROLLO EDUCATIVO S.A. DE C.V."
    ]

    activos[col_institucion] = activos[col_institucion].astype(str).str.strip()
    activos_filtrado = activos[~activos[col_institucion].isin(instituciones_a_borrar)].copy()

    # Detectar "No. Póliza"
    col_poliza = None
    for c in activos_filtrado.columns:
        if c.strip().lower() == "no. póliza".lower():
            col_poliza = c
            break

    if col_poliza:
        activos_filtrado[col_poliza] = pd.to_numeric(activos_filtrado[col_poliza], errors="coerce")
    else:
        raise ValueError("⚠ No encontré 'No. Póliza' en Activos_filtrado")

    # ==============================
    # PASO 2 — AGREGAR COLUMNAS NUEVAS
    # ==============================
    for col in ["Reporte", "Activos GNP"]:
        if col not in activos_filtrado.columns:
            activos_filtrado[col] = None

    # ==============================
    # PASO 3 — LLENAR Reporte y Activos GNP
    # ==============================
    col_ref_nomina = None
    for c in nomina.columns:
        if c.strip().lower() == "nº ref.externo".lower():
            col_ref_nomina = c
            break
    if col_ref_nomina is None:
        raise ValueError("No encontré 'Nº ref.externo' en nomina")

    col_cdnum_parque = None
    for c in parque_vigentes.columns:
        if c.strip().lower() == "cdnumpol".lower():
            col_cdnum_parque = c
            break
    if col_cdnum_parque is None:
        raise ValueError("No encontré 'CDNUMPOL' en Parque.xlsx hoja Vigentes")

    llave_activos = activos_filtrado[col_poliza].astype(str).str.strip()
    llave_nomina  = nomina[col_ref_nomina].astype(str).str.strip()
    llave_parque  = parque_vigentes[col_cdnum_parque].astype(str).str.strip()

    valores_nomina = set(llave_nomina.dropna())
    valores_parque = set(llave_parque.dropna())

    # Reporte = está en nómina
    activos_filtrado["Reporte"] = llave_activos.where(llave_activos.isin(valores_nomina), pd.NA)
    # Activos GNP = está en parque
    activos_filtrado["Activos GNP"] = llave_activos.where(llave_activos.isin(valores_parque), pd.NA)

    # ==============================
    # PASO 4 — FILTRAR ÚLTIMO MES COMPLETO (Fecha de Alta)
    # ==============================
    col_fecha_alta = None
    for c in activos_filtrado.columns:
        if c.strip().lower() == "fecha de alta":
            col_fecha_alta = c
            break
    if col_fecha_alta is None:
        raise ValueError("No encontré la columna 'Fecha de Alta' en Activos_filtrado")

    activos_filtrado[col_fecha_alta] = pd.to_datetime(
        activos_filtrado[col_fecha_alta],
        errors="coerce",
        dayfirst=True
    )

    hoy = pd.Timestamp.today().normalize()
    primer_dia_mes_actual = hoy.replace(day=1)
    primer_dia_hace_un_mes = primer_dia_mes_actual - DateOffset(months=1)

    activos_filtrado = activos_filtrado[
        activos_filtrado[col_fecha_alta] >= primer_dia_hace_un_mes
    ]

    # ==============================
    # PASO 5 — FILTRO PARA TEMPLATE (SOLO DONDE REPORTE ES NULO)
    # ==============================
    df_template_src = activos_filtrado[activos_filtrado["Reporte"].isna()].copy()

    # ==============================
    # LIMPIAR NÚMERO DE NÓMINA → NÚMERO DE PERSONAL (SOLO PARA TEMPLATE)
    # ==============================
    col_num_nomina = None
    for c in df_template_src.columns:
        if "nomina" in c.lower() or "nómina" in c.lower():
            col_num_nomina = c
            break
    if col_num_nomina is None:
        raise ValueError("❌ No encontré la columna de número de nómina en Activos_filtrado")

    df_template_src["NumeroPersonalLimpio"] = (
        df_template_src[col_num_nomina]
            .astype(str)
            .str.strip()
            .str.upper()
            .str.replace("L", "", regex=False)
    )
    df_template_src["NumeroPersonalLimpio"] = pd.to_numeric(
        df_template_src["NumeroPersonalLimpio"],
        errors="coerce"
    )

    # ==============================
    # PASO 6 — LLENAR TEMPLATE DE ALTAS (EN MEMORIA)
    # ==============================
    template_final = pd.DataFrame(
        index=range(len(df_template_src)),
        columns=template_altas.columns
    )

    # Número de personal
    if "Número de personal" in template_final.columns:
        template_final["Número de personal"] = df_template_src["NumeroPersonalLimpio"].values

    # Valores fijos
    template_final["Tipo de carga"] = "C"
    template_final["Fin de validez"] = "31.12.9999"
    template_final["Cc-nómina"] = "3353"
    template_final["Condición préstamo SE01-QU02"] = "02"
    template_final["Vía de Pago"] = "0100"
    template_final["Texto"] = "Seguro de Automóvil altas NOM"
    template_final["Subdivisión"] = "0001"

    # ==============================
    # N° referencia externo = No. Póliza (de los que TENÍAN Reporte nulo)
    # ==============================
    col_template_ref_ext = None
    for c in template_final.columns:
        if "referencia" in c.lower() and "externo" in c.lower():
            col_template_ref_ext = c
            break

    if col_template_ref_ext is not None and col_poliza is not None:
        template_final[col_template_ref_ext] = (
            df_template_src[col_poliza]
            .astype("Int64")
            .astype(str)
            .str.replace("<NA>", "", regex=False)
            .values
        )

    # ==============================
    # FECHAS CON NUEVA LÓGICA DE QUINCENA
    # ==============================
    hoy2 = pd.Timestamp.today()

    # Si es del 1 al 15 -> 16 del mismo mes
    # Si es del 16 al final -> 1 del siguiente mes
    if hoy2.day <= 15:
        fecha_quincena = hoy2.replace(day=16)
    else:
        # pasar al 1 del siguiente mes
        if hoy2.month == 12:
            fecha_quincena = hoy2.replace(year=hoy2.year + 1, month=1, day=1)
        else:
            fecha_quincena = hoy2.replace(month=hoy2.month + 1, day=1)

    fecha_quincena_str = fecha_quincena.strftime("%d.%m.%Y")

    if "Inicio de la validez" in template_final.columns:
        template_final["Inicio de la validez"] = fecha_quincena_str

    columnas_quincena = [
        "Fecha de la autorización",
        "Inicio de Amortización",
        "Fecha de Pago"
    ]
    for col in columnas_quincena:
        if col in template_final.columns:
            template_final[col] = fecha_quincena_str

    # ==============================
    # IMPORTE DE PRÉSTAMO AUTORIZADO
    # ==============================
    col_importe_origen = None
    for c in df_template_src.columns:
        if c.strip().lower() == "precio a fin de vigencia".lower():
            col_importe_origen = c
            break
    if col_importe_origen is None:
        raise ValueError("No encontré la columna 'Precio a Fin de Vigencia' en Activos_filtrado")

    col_importe_destino = "Importe de préstamo autorizado"
    if col_importe_destino in template_final.columns:
        template_final[col_importe_destino] = df_template_src[col_importe_origen].values

    # ==============================
    # AMORTIZACIÓN
    # ==============================
    mes_inicio = fecha_quincena.month
    dia_inicio = fecha_quincena.day

    quincenas_restantes = 0
    for mes in range(mes_inicio, 13):
        if mes == mes_inicio:
            if dia_inicio == 1:
                quincenas_restantes += 2
            else:
                quincenas_restantes += 1
        else:
            quincenas_restantes += 2

    col_amortizacion = "Amortización"
    if (col_importe_destino in template_final.columns) and (col_amortizacion in template_final.columns):
        importe_num = pd.to_numeric(template_final[col_importe_destino], errors="coerce")
        if quincenas_restantes > 0:
            template_final[col_amortizacion] = (importe_num / quincenas_restantes).round(2)

    # ==============================
    # ACTIVOS_FILTRADO DE SALIDA (REPORTE Y ACTIVOS GNP CON N/A)
    # ==============================
    activos_salida = activos_filtrado.copy()
    activos_salida["Reporte"] = activos_salida["Reporte"].fillna("N/A")
    activos_salida["Activos GNP"] = activos_salida["Activos GNP"].fillna("N/A")
    activos_salida[col_fecha_alta] = pd.to_datetime(
        activos_salida[col_fecha_alta], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    return template_final, activos_salida


# =========================================
# DESCARGA A EXCEL (CON FECHA EN EL NOMBRE)
# =========================================

def df_to_excel_download(df, filename, label=None):
    # fecha de hoy en formato AAAA-MM-DD
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    nombre_final = f"{fecha_hoy}_{filename}"

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    st.download_button(
        label=label or f"📥 Descargar {nombre_final}",
        data=buffer,
        file_name=nombre_final,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================================
# APP STREAMLIT
# =========================================

st.set_page_config(page_title="Raspberry – Altas y Bajas", layout="wide")

st.title("Reportes Raspberry – Altas y Bajas")

tab_bajas, tab_altas = st.tabs(["🔻 Reportes de Bajas", "🔺 Reportes de Altas"])

# ---------------- TAB BAJAS ----------------
with tab_bajas:
    st.subheader("Consolidado de Bajas")

    parque_file = st.file_uploader("Parque vehicular", type=["xlsx", "xls"], key="parque_bajas")
    cancelacion_file = st.file_uploader("Cancelación", type=["xlsx", "xls"], key="cancelacion_bajas")
    cancelado_file = st.file_uploader("Cancelado", type=["xlsx", "xls"], key="cancelado_bajas")
    nomina_file = st.file_uploader("Desectos o nomina", type=["xlsx", "xls"], key="nomina_bajas")

    # 🕒 Momento en que ya se tienen todos los archivos cargados
    if all([parque_file, cancelacion_file, cancelado_file, nomina_file]):
        if "bajas_upload_time" not in st.session_state:
            st.session_state.bajas_upload_time = datetime.now()

        st.info(
            f"📂 Archivos de BAJAS cargados el: "
            f"{st.session_state.bajas_upload_time:%Y-%m-%d %H:%M:%S}"
        )

        if st.button("Procesar bajas"):
            start_time = datetime.now()
            st.session_state.bajas_start_time = start_time

            parque_df = pd.read_excel(parque_file, sheet_name="Anuladas")
            cancelacion_df = pd.read_excel(cancelacion_file)
            cancelado_df = pd.read_excel(cancelado_file)
            nomina_df = pd.read_excel(nomina_file)

            for df in [parque_df, cancelacion_df, cancelado_df, nomina_df]:
                df.columns = df.columns.str.strip()

            consolidado = procesar_bajas(parque_df, cancelacion_df, cancelado_df, nomina_df)

            end_time = datetime.now()
            st.session_state.bajas_end_time = end_time

            st.success("Consolidado de bajas generado correctamente.")
            st.write(f"⏱ Inicio del proceso: {start_time:%Y-%m-%d %H:%M:%S}")
            st.write(f"✅ Fin del proceso: {end_time:%Y-%m-%d %H:%M:%S}")
            st.write(f"⌛ Duración: {(end_time - start_time).total_seconds():.1f} segundos")

            df_to_excel_download(consolidado, "consolidado_de_bajas.xlsx", label="📥 Descargar bajas")
    else:
        st.info("📂 Sube todos los archivos para poder generar el consolidado de bajas.")

# ---------------- TAB ALTAS ----------------
with tab_altas:
    st.subheader("Template de Altas")

    parque_v_file = st.file_uploader("Parque vehicular ", type=["xlsx", "xls"], key="parque_altas")
    activos_file = st.file_uploader("Activos", type=["xlsx", "xls"], key="activos_altas")
    nomina_altas_file = st.file_uploader("Desectos o Nominas", type=["xlsx", "xls"], key="nomina_altas")
    template_file = st.file_uploader("Altas (Template)", type=["xlsx", "xls"], key="template_altas")

    if all([parque_v_file, activos_file, nomina_altas_file, template_file]):
        if "altas_upload_time" not in st.session_state:
            st.session_state.altas_upload_time = datetime.now()

        st.info(
            f"📂 Archivos de ALTAS cargados el: "
            f"{st.session_state.altas_upload_time:%Y-%m-%d %H:%M:%S}"
        )

        if st.button("Procesar altas"):
            start_time = datetime.now()
            st.session_state.altas_start_time = start_time

            parque_v_df = pd.read_excel(parque_v_file, sheet_name="Vigentes")
            activos_df = pd.read_excel(activos_file)
            nomina_altas_df = pd.read_excel(nomina_altas_file)
            template_df = pd.read_excel(template_file)

            for df in [parque_v_df, activos_df, nomina_altas_df, template_df]:
                df.columns = df.columns.str.strip()

            template_final_df, activos_salida_df = procesar_altas(
                parque_v_df, activos_df, nomina_altas_df, template_df
            )

            end_time = datetime.now()
            st.session_state.altas_end_time = end_time

            st.success("Template de Altas generado correctamente.")
            st.write(f"⏱ Inicio del proceso: {start_time:%Y-%m-%d %H:%M:%S}")
            st.write(f"✅ Fin del proceso: {end_time:%Y-%m-%d %H:%M:%S}")
            st.write(f"⌛ Duración: {(end_time - start_time).total_seconds():.1f} segundos")

            df_to_excel_download(
                template_final_df,
                "Template_de_Altas_generado.xlsx",
                label="📥 Descargar Template de Altas"
            )
            df_to_excel_download(
                activos_salida_df,
                "Activos_filtrados_altas.xlsx",
                label="📥 Descargar Activos"
            )
    else:
        st.info("📂 Sube todos los archivos para poder generar el reporte de Altas.")
