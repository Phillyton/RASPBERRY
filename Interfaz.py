import streamlit as st
import pandas as pd
from pandas import DateOffset
from io import BytesIO
from datetime import datetime
import pytz

# =========================================
# FUNCIÓN GLOBAL DE DESCARGA CON FECHA MX
# =========================================

def df_to_excel_download(df, filename, label=None):

    tz = pytz.timezone("America/Mexico_City")  # ← fecha correcta MX
    today_str = datetime.now(tz).strftime("%Y-%m-%d")

    # agrega fecha ANTES del nombre
    if "." in filename:
        name, ext = filename.rsplit(".", 1)
        file_name = f"{today_str}_{name}.{ext}"
    else:
        file_name = f"{today_str}_{filename}.xlsx"

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)

    st.download_button(
        label=label or f"📥 Descargar {file_name}",
        data=buffer,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================================
# FUNCIONES DE PROCESO (BAJAS Y ALTAS)
# =========================================

def procesar_bajas(parque, cancelacion, cancelado, nomina):
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

    COL_LLAVE_NOMINA    = "Nº ref.externo"
    COL_LLAVE_CANCELADO = "No. Póliza"
    COL_FECHA_BAJA_OP   = "Fecha de Baja Operativa"

    for df, col in [(nomina, COL_LLAVE_NOMINA), (cancelado, COL_LLAVE_CANCELADO)]:
        df[col] = df[col].astype(str).str.strip()

    cancelado_tmp = cancelado.copy()
    cancelado_tmp[COL_FECHA_BAJA_OP] = pd.to_datetime(cancelado_tmp[COL_FECHA_BAJA_OP], errors="coerce")

    cancelado_tmp = (
        cancelado_tmp.sort_values(COL_FECHA_BAJA_OP)
        .drop_duplicates(subset=COL_LLAVE_CANCELADO, keep="first")
        .set_index(COL_LLAVE_CANCELADO)
    )

    serie_poliza = pd.Series(cancelado_tmp.index, index=cancelado_tmp.index)

    nomina["Cancelado"] = nomina[COL_LLAVE_NOMINA].map(serie_poliza)
    nomina["Fecha de cancelado"] = nomina[COL_LLAVE_NOMINA].map(cancelado_tmp[COL_FECHA_BAJA_OP])

    df_nomina_cancelaciones_antes = nomina.copy()
    df_nomina_cancelaciones_antes = df_nomina_cancelaciones_antes.drop(
        columns=["Cancelado", "Fecha de cancelado"], errors="ignore"
    )

    nomina["Cancelado"] = nomina["Cancelado"].replace("", pd.NA)
    nomina = nomina[nomina["Cancelado"].notna()]

    columnas_valores = ["SaldoIni", "DesctPer", "SaldoFin"]

    for col in columnas_valores:
        if col in nomina.columns:
            nomina[col] = pd.to_numeric(nomina[col], errors="coerce")

    nomina = nomina[
        (nomina["SaldoIni"] > 0) &
        (nomina["DesctPer"] > 0) &
        (nomina["SaldoFin"] > 0)
    ]

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
            nomina[col] = pd.to_datetime(nomina[col], errors="coerce").dt.strftime("%d/%m/%Y")

    columnas_vacias = [col for col in nomina.columns if nomina[col].isna().all()]
    if columnas_vacias:
        nomina.drop(columns=columnas_vacias, inplace=True)

    # CANCELACIÓN
    df_cancelacion_nomina = df_nomina_cancelaciones_antes.copy()

    COL_LLAVE_NOMINA_CANC  = "Nº ref.externo"
    COL_LLAVE_CANCELACION  = "No. Póliza"
    COL_FECHA_BAJA_OP_CANC = "Fecha de Baja Operativa"

    for df_tmp, col in [(df_cancelacion_nomina, COL_LLAVE_NOMINA_CANC),
                        (cancelacion, COL_LLAVE_CANCELACION)]:
        df_tmp[col] = df_tmp[col].astype(str).str.strip()

    cancelacion_tmp = cancelacion.copy()
    cancelacion_tmp[COL_FECHA_BAJA_OP_CANC] = pd.to_datetime(cancelacion_tmp[COL_FECHA_BAJA_OP_CANC],
                                                             errors="coerce")

    cancelacion_tmp = (
        cancelacion_tmp.sort_values(COL_FECHA_BAJA_OP_CANC)
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

    df_cancelacion_nomina = df_cancelacion_nomina[df_cancelacion_nomina["Cancelación"].notna()]

    # ANULADOS GNP
    df_anulados_nomina = df_nomina_cancelaciones_antes.copy()

    COL_LLAVE_NOMINA_ANU = "Nº ref.externo"
    COL_LLAVE_PARQUE_ANU = "CDNUMPOL"

    for df_tmp, col in [(df_anulados_nomina, COL_LLAVE_NOMINA_ANU),
                        (parque, COL_LLAVE_PARQUE_ANU)]:
        df_tmp[col] = df_tmp[col].astype(str).str.strip()

    parque_tmp = parque.copy()
    parque_tmp = parque_tmp.drop_duplicates(subset=COL_LLAVE_PARQUE_ANU, keep="first").set_index(COL_LLAVE_PARQUE_ANU)

    df_anulados_nomina["Anulados GNP"] = df_anulados_nomina[COL_LLAVE_NOMINA_ANU].map(parque_tmp.index)
    df_anulados_nomina = df_anulados_nomina[df_anulados_nomina["Anulados GNP"].notna()]

    # CONSOLIDADO FINAL
    df_cancelado_consol = nomina.copy()
    df_cancelacion_consol = df_cancelacion_nomina.copy()
    df_anulados_consol = df_anulados_nomina.copy()

    columnas_originales = df_nomina_cancelaciones_antes.columns.tolist()

    df_cancelado_consol = df_cancelado_consol.reindex(columns=columnas_originales)
    df_cancelacion_consol = df_cancelacion_consol.reindex(columns=columnas_originales)
    df_anulados_consol = df_anulados_consol.reindex(columns=columnas_originales)

    df_cancelado_consol["Tipo_baja"] = "Cancelado"
    df_cancelacion_consol["Tipo_baja"] = "Cancelación"
    df_anulados_consol["Tipo_baja"] = "Anulado GNP"

    consolidado_bajas = pd.concat(
        [df_cancelado_consol, df_cancelacion_consol, df_anulados_consol],
        ignore_index=True
    )

    cols_drop = ["Cancelado", "Fecha de cancelado", "Cancelación", "Fecha de Cancelación", "Anulados GNP"]
    consolidado_bajas = consolidado_bajas.drop(columns=cols_drop, errors="ignore")

    return consolidado_bajas


# =========================================
# ALTAS
# =========================================

def procesar_altas(parque_vigentes, activos, nomina, template_altas):
    # ... (tu código largo de ALTAS se mantiene igual, no lo modifico aquí)
    # Para no hacer este mensaje de 5000 líneas, mantengo intacto tu bloque ALTAS

    # Solo agrego devolución:
    return template_final, activos_salida


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

    if all([parque_file, cancelacion_file, cancelado_file, nomina_file]):
        if st.button("Procesar bajas"):
            parque_df = pd.read_excel(parque_file, sheet_name="Anuladas")
            cancelacion_df = pd.read_excel(cancelacion_file)
            cancelado_df = pd.read_excel(cancelado_file)
            nomina_df = pd.read_excel(nomina_file)

            for df in [parque_df, cancelacion_df, cancelado_df, nomina_df]:
                df.columns = df.columns.str.strip()

            consolidado = procesar_bajas(parque_df, cancelacion_df, cancelado_df, nomina_df)
            st.success("Consolidado de bajas generado correctamente.")

            df_to_excel_download(
                consolidado,
                "Consolidado_de_Bajas.xlsx",
                label="Descargar Consolidado de Bajas"
            )

# ---------------- TAB ALTAS ----------------
with tab_altas:
    st.subheader("Template de Altas")

    parque_v_file = st.file_uploader("Parque vehicular", type=["xlsx", "xls"], key="parque_altas")
    activos_file = st.file_uploader("Activos", type=["xlsx", "xls"], key="activos_altas")
    nomina_altas_file = st.file_uploader("Desectos o Nominas", type=["xlsx", "xls"], key="nomina_altas")
    template_file = st.file_uploader("Altas (Template)", type=["xlsx", "xls"], key="template_altas")

    if all([parque_v_file, activos_file, nomina_altas_file, template_file]):
        if st.button("Procesar altas"):
            parque_v_df = pd.read_excel(parque_v_file, sheet_name="Vigentes")
            activos_df = pd.read_excel(activos_file)
            nomina_altas_df = pd.read_excel(nomina_altas_file)
            template_df = pd.read_excel(template_file)

            for df in [parque_v_df, activos_df, nomina_altas_df, template_df]:
                df.columns = df.columns.str.strip()

            template_final_df, activos_salida_df = procesar_altas(
                parque_v_df, activos_df, nomina_altas_df, template_df
            )

            st.success("Template de Altas generado correctamente.")

            df_to_excel_download(
                template_final_df,
                "Template_de_Altas_generado.xlsx",
                label="Descargar Template de Altas"
            )

            df_to_excel_download(
                activos_salida_df,
                "Activos_filtrados_altas.xlsx",
                label="Descargar Activos Altas"
            )
