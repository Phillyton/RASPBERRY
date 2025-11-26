import streamlit as st
import pandas as pd
from pandas import DateOffset
from io import BytesIO

# =========================================
# FUNCIONES DE UTILIDAD
# =========================================

def df_to_excel_download(df, filename):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    st.download_button(
        label=f"📥 Descargar {filename}",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# =========================================
# PROCESO DE BAJAS 
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
    cancelado_tmp[COL_FECHA_BAJA_OP] = pd.to_datetime(
        cancelado_tmp[COL_FECHA_BAJA_OP], errors="coerce"
    )

    cancelado_tmp = (
        cancelado_tmp
        .sort_values(COL_FECHA_BAJA_OP)
        .drop_duplicates(subset=COL_LLAVE_CANCELADO, keep="first")
        .set_index(COL_LLAVE_CANCELADO)
    )

    serie_poliza = pd.Series(cancelado_tmp.index, index=cancelado_tmp.index)

    nomina["Cancelado"] = nomina[COL_LLAVE_NOMINA].map(serie_poliza)
    nomina["Fecha de cancelado"] = nomina[COL_LLAVE_NOMINA].map(
        cancelado_tmp[COL_FECHA_BAJA_OP]
    )

    df_nomina_cancelaciones_antes_de_eliminar = nomina.copy()
    df_nomina_cancelaciones_antes_de_eliminar = df_nomina_cancelaciones_antes_de_eliminar.drop(
        columns=["Cancelado", "Fecha de cancelado"],
        errors="ignore"
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
            nomina[col] = pd.to_datetime(
                nomina[col], errors="coerce"
            ).dt.strftime("%d/%m/%Y")

    columnas_vacias = [col for col in nomina.columns if nomina[col].isna().all()]
    if columnas_vacias:
        nomina.drop(columns=columnas_vacias, inplace=True)

    # ---------- Cancelación ----------
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

    # ---------- Anulados GNP ----------
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

    df_anulados_nomina = df_anulados_nomina.drop(
        columns=["Cancelación", "Fecha de Cancelación"],
        errors="ignore"
    )

    # ---------- Consolidado ----------
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
    # 1.1 Columna Institución
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

    # 1.2 Detectar "No. Póliza"
    col_poliza = None
    for c in activos_filtrado.columns:
        if c.strip().lower() == "no. póliza".lower():
            col_poliza = c
            break
    if col_poliza is None:
        raise ValueError("⚠ No encontré 'No. Póliza' en Activos_filtrado")

    activos_filtrado[col_poliza] = pd.to_numeric(activos_filtrado[col_poliza], errors="coerce")

    # ==============================
    # PASO 2 — AGREGAR COLUMNAS NUEVAS
    # ==============================
    for col in ["Reporte", "Activos GNP"]:
        if col not in activos_filtrado.columns:
            activos_filtrado[col] = None

    # ==============================
    # PASO 3 — LLENAR Reporte y Activos GNP (SIN FILTRAR)
    # ==============================
    # 3.1 'Nº ref.externo' en nómina
    col_ref_nomina = None
    for c in nomina.columns:
        if c.strip().lower() == "nº ref.externo".lower():
            col_ref_nomina = c
            break
    if col_ref_nomina is None:
        raise ValueError("No encontré 'Nº ref.externo' en nomina")

    # 3.2 'CDNUMPOL' en parque vigentes
    col_cdnum_parque = None
    for c in parque_vigentes.columns:
        if c.strip().lower() == "cdnumpol".lower():
            col_cdnum_parque = c
            break
    if col_cdnum_parque is None:
        raise ValueError("No encontré 'CDNUMPOL' en Parque.xlsx hoja Vigentes")

    # 3.3 Normalizar llaves
    llave_activos = activos_filtrado[col_poliza].astype(str).str.strip()
    llave_nomina  = nomina[col_ref_nomina].astype(str).str.strip()
    llave_parque  = parque_vigentes[col_cdnum_parque].astype(str).str.strip()

    valores_nomina  = set(llave_nomina.dropna())
    valores_parque  = set(llave_parque.dropna())

    # Llenar (solo información, ya NO filtramos por esto)
    activos_filtrado["Reporte"]     = llave_activos.where(llave_activos.isin(valores_nomina), pd.NA)
    # Por lo que pediste: Activos GNP fijo en "N/A"
    activos_filtrado["Activos GNP"] = "N/A"

    # ==============================
    # PASO 4 — FILTRAR ÚLTIMOS 2 MESES COMPLETOS (Fecha de Alta)
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
    primer_dia_hace_dos_meses = primer_dia_mes_actual - DateOffset(months=2)

    activos_filtrado = activos_filtrado[
        activos_filtrado[col_fecha_alta] >= primer_dia_hace_dos_meses
    ]

    # Para que se vea bonito en el Excel de salida de activos
    activos_filtrado[col_fecha_alta] = activos_filtrado[col_fecha_alta].dt.strftime("%d/%m/%Y")

    # ==============================
    # PASO 5 — LIMPIAR NÚMERO DE NÓMINA → NÚMERO DE PERSONAL
    # ==============================
    col_num_nomina = None
    for c in activos_filtrado.columns:
        if "nomina" in c.lower() or "nómina" in c.lower():
            col_num_nomina = c
            break
    if col_num_nomina is None:
        raise ValueError("❌ No encontré la columna de número de nómina en Activos_filtrado")

    activos_filtrado["NumeroPersonalLimpio"] = (
        activos_filtrado[col_num_nomina]
            .astype(str)
            .str.strip()
            .str.upper()
            .str.replace("L", "", regex=False)
    )
    activos_filtrado["NumeroPersonalLimpio"] = pd.to_numeric(
        activos_filtrado["NumeroPersonalLimpio"],
        errors="coerce"
    )

    # ==============================
    # PASO 6 — CREAR TEMPLATE DE ALTAS CON MISMO NÚMERO DE FILAS
    # ==============================
    # En lugar de recortar el template original, creamos uno nuevo con
    # las MISMAS COLUMNAS pero tantas filas como registros tenga activos_filtrado
    template_final = pd.DataFrame(
        index=range(len(activos_filtrado)),
        columns=template_altas.columns
    )

    # 6.1 Número de personal
    if "Número de personal" in template_final.columns:
        template_final["Número de personal"] = activos_filtrado["NumeroPersonalLimpio"].values

    # 6.2 Valores fijos
    template_final["Tipo de carga"]                  = "C"
    template_final["Fin de validez"]                 = "31.12.9999"
    template_final["Cc-nómina"]                      = "3353"
    template_final["Condición préstamo SE01-QU02"]   = "02"
    template_final["Vía de Pago"]                    = "0100"
    template_final["Texto"]                          = "Seguro de Automóvil altas NOM"
    template_final["Subdivisión"]                    = "0001"

    # ==============================
    # PASO 7 — N° referencia externo (desde No. Póliza)
    # ==============================
    col_template_ref_ext = None
    for c in template_final.columns:
        if "referencia" in c.lower() and "externo" in c.lower():
            col_template_ref_ext = c
            break

    if col_template_ref_ext is not None:
        template_final[col_template_ref_ext] = (
            activos_filtrado[col_poliza]
            .astype("Int64")
            .astype(str)
            .str.replace("<NA>", "", regex=False)
            .values
        )

    # ==============================
    # PASO 8 — FECHAS DE QUINCENA
    # ==============================
    hoy2 = pd.Timestamp.today()
    if hoy2.day <= 15:
        fecha_quincena = hoy2.replace(day=1)
    else:
        fecha_quincena = hoy2.replace(day=16)

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
    # PASO 9 — IMPORTE DE PRÉSTAMO AUTORIZADO
    # ==============================
    col_importe_origen = None
    for c in activos_filtrado.columns:
        if c.strip().lower() == "precio a fin de vigencia".lower():
            col_importe_origen = c
            break
    if col_importe_origen is None:
        raise ValueError("No encontré la columna 'Precio a Fin de Vigencia' en Activos_filtrado")

    col_importe_destino = "Importe de préstamo autorizado"
    if col_importe_destino in template_final.columns:
        template_final[col_importe_destino] = activos_filtrado[col_importe_origen].values

    # ==============================
    # PASO 10 — AMORTIZACIÓN
    # ==============================
    mes_inicio = fecha_quincena.month
    dia_inicio = fecha_quincena.day

    quincenas_restantes = 0
    for mes in range(mes_inicio, 13):
        if mes == mes_inicio:
            quincenas_restantes += 2 if dia_inicio == 1 else 1
        else:
            quincenas_restantes += 2

    col_amortizacion = "Amortización"
    if (col_importe_destino in template_final.columns) and (col_amortizacion in template_final.columns):
        importe_num = pd.to_numeric(template_final[col_importe_destino], errors="coerce")
        if quincenas_restantes > 0:
            template_final[col_amortizacion] = (importe_num / quincenas_restantes).round(2)

    # ==============================
    # PASO 11 — PREPARAR SALIDA DE ACTIVOS DEPURADOS
    # ==============================
    activos_salida = activos_filtrado.copy()
    # No queremos mandar NumeroPersonalLimpio en el Excel depurado
    activos_salida = activos_salida.drop(columns=["NumeroPersonalLimpio"], errors="ignore")

    return template_final, activos_salida
