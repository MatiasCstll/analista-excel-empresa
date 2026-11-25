import io
import json
import re
import time
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfReader

st.set_page_config(
    page_title="Analista Excel IA",
    page_icon="📊",
    layout="wide",
)


# ---------------------------------------------------------
# 1. CONTROL DE ACCESO (USUARIOS DEFINIDOS EN SECRETS)
# ---------------------------------------------------------
def verificar_login():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False
        st.session_state.usuario_actual = None

    if st.session_state.autenticado:
        return True

    st.title("🔐 Acceso al Analista Excel Corporativo")
    st.caption("Ingresá con el usuario habilitado por el administrador.")

    with st.form("form_login"):
        usuario = st.text_input("Usuario").strip().lower()
        password = st.text_input("Contraseña", type="password")
        enviar = st.form_submit_button("Ingresar", use_container_width=True)

    if enviar:
        usuarios_validos = st.secrets.get("usuarios", {})
        if usuario in usuarios_validos and str(usuarios_validos[usuario]) == password:
            st.session_state.autenticado = True
            st.session_state.usuario_actual = usuario
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos, o tu acceso fue deshabilitado.")

    return False


# ---------------------------------------------------------
# 2. NORMALIZACIÓN INTELIGENTE DE TABLAS Y TEXTOS
# ---------------------------------------------------------
def normalizar_numero(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    s_clean = re.sub(r"[$€£¥\s]", "", s)
    if not re.match(r"^-?[\d.,]+$", s_clean):
        return val
    # Formato argentino/europeo: 1.250,50 -> 1250.50
    if "," in s_clean and "." in s_clean:
        if s_clean.rfind(",") > s_clean.rfind("."):
            s_clean = s_clean.replace(".", "").replace(",", ".")
        else:
            s_clean = s_clean.replace(",", "")
    elif "," in s_clean:
        partes = s_clean.split(",")
        if len(partes) == 2 and len(partes[1]) in (1, 2):
            s_clean = s_clean.replace(",", ".")
        else:
            s_clean = s_clean.replace(",", "")
    try:
        return float(s_clean)
    except ValueError:
        return val


def limpiar_dataframe_erp(df_raw: pd.DataFrame) -> pd.DataFrame:
    if df_raw.empty:
        return df_raw

    # Detectar si hay filas de título/logo arriba de los encabezados reales
    cols_unnamed = sum(1 for c in df_raw.columns if str(c).startswith("Unnamed:"))
    if cols_unnamed >= max(1, len(df_raw.columns) // 2):
        best_row = 0
        max_non_null = -1
        for i in range(min(15, len(df_raw))):
            non_nulls = df_raw.iloc[i].dropna()
            str_count = sum(1 for x in non_nulls if isinstance(x, str) and len(x.strip()) > 0)
            if str_count > max_non_null and str_count >= 2:
                max_non_null = str_count
                best_row = i
        df = df_raw.iloc[best_row + 1 :].copy()
        df.columns = [
            str(c).strip() if pd.notna(c) and str(c).strip() else f"Columna_{idx + 1}"
            for idx, c in enumerate(df_raw.iloc[best_row])
        ]
    else:
        df = df_raw.copy()
        df.columns = [str(c).strip() for c in df.columns]

    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    # Normalizar columnas numéricas y fechas
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d").fillna("")
            continue
        convertidos = df[col].apply(normalizar_numero)
        no_nulos = [x for x in convertidos if x is not None and str(x).strip() != ""]
        num_count = sum(1 for x in no_nulos if isinstance(x, (int, float)))
        if no_nulos and (num_count / len(no_nulos)) >= 0.7:
            df[col] = [x if isinstance(x, (int, float)) else 0.0 for x in convertidos]
        else:
            df[col] = df[col].fillna("").astype(str).str.strip()

    return df


def preparar_entrada(texto_pegado: str, archivo):
    """
    Devuelve:
    - modo: 'tabular' (si subió Excel/CSV limpio) o 'texto' (si pegó texto o subió PDF/TXT/JSON)
    - df_tabular: DataFrame completo limpio (solo en modo 'tabular')
    - contenido_para_ia: Texto optimizado para enviar a Gemini
    - respaldo_crudo: Texto de respaldo para la pestaña 'Crudo Original'
    """
    if archivo is not None and not (texto_pegado and texto_pegado.strip()):
        nombre = archivo.name.lower()
        contenido_bytes = archivo.read()

        if nombre.endswith((".xlsx", ".xls")):
            hojas = pd.read_excel(io.BytesIO(contenido_bytes), sheet_name=None)
            mejor_hoja = max(
                hojas.items(),
                key=lambda item: item[1].dropna(how="all").shape[0] * max(item[1].shape[1], 1),
            )
            nombre_hoja, df_raw = mejor_hoja
            df_limpio = limpiar_dataframe_erp(df_raw)

            if not df_limpio.empty and len(df_limpio.columns) >= 2:
                resumen_cols = []
                for c in df_limpio.columns:
                    if pd.api.types.is_numeric_dtype(df_limpio[c]):
                        resumen_cols.append(
                            f"- {c} (Numérica): Suma={df_limpio[c].sum():,.2f}, "
                            f"Promedio={df_limpio[c].mean():,.2f}, Máx={df_limpio[c].max():,.2f}"
                        )
                    else:
                        top_vals = df_limpio[c].value_counts().head(5).to_dict()
                        resumen_cols.append(f"- {c} (Texto/Categoría): Top valores={top_vals}")

                muestra_csv = df_limpio.head(20).to_csv(index=False)
                contenido_ia = (
                    f"ARCHIVO EXCEL: {archivo.name} (Hoja: {nombre_hoja})\n"
                    f"TOTAL DE FILAS REALES: {len(df_limpio)}\n\n"
                    f"ESTADÍSTICAS EXACTAS CALCULADAS SOBRE EL 100% DE LOS DATOS:\n"
                    + "\n".join(resumen_cols)
                    + f"\n\nMUESTRA DE LAS PRIMERAS 20 FILAS:\n{muestra_csv}"
                )
                respaldo = df_limpio.head(500).to_csv(index=False)
                return "tabular", df_limpio, contenido_ia, respaldo

        elif nombre.endswith(".csv"):
            df_raw = pd.DataFrame()
            for sep in (",", ";", "\t"):
                try:
                    df_raw = pd.read_csv(io.BytesIO(contenido_bytes), sep=sep, encoding="utf-8")
                    if df_raw.shape[1] >= 2:
                        break
                except Exception:
                    try:
                        df_raw = pd.read_csv(io.BytesIO(contenido_bytes), sep=sep, encoding="latin-1")
                        if df_raw.shape[1] >= 2:
                            break
                    except Exception:
                        df_raw = pd.DataFrame()
            df_limpio = limpiar_dataframe_erp(df_raw)
            if not df_limpio.empty and len(df_limpio.columns) >= 2:
                resumen_cols = []
                for c in df_limpio.columns:
                    if pd.api.types.is_numeric_dtype(df_limpio[c]):
                        resumen_cols.append(
                            f"- {c} (Numérica): Suma={df_limpio[c].sum():,.2f}, "
                            f"Promedio={df_limpio[c].mean():,.2f}"
                        )
                    else:
                        top_vals = df_limpio[c].value_counts().head(5).to_dict()
                        resumen_cols.append(f"- {c} (Texto/Categoría): Top valores={top_vals}")
                contenido_ia = (
                    f"ARCHIVO CSV: {archivo.name} | TOTAL FILAS: {len(df_limpio)}\n"
                    + "\n".join(resumen_cols)
                    + f"\n\nMUESTRA (20 FILAS):\n{df_limpio.head(20).to_csv(index=False)}"
                )
                return "tabular", df_limpio, contenido_ia, df_limpio.head(500).to_csv(index=False)

        # Si era PDF, TXT, JSON o un Excel desestructurado, volver a modo texto
        archivo.seek(0)

    # Modo texto / PDF / Mixto
    bloques = []
    if texto_pegado and texto_pegado.strip():
        bloques.append("=== TEXTO PEGADO ===\n" + texto_pegado.strip())

    if archivo is not None:
        nombre = archivo.name.lower()
        contenido_bytes = archivo.read()
        if nombre.endswith((".csv", ".txt", ".json", ".tsv", ".log")):
            for enc in ("utf-8", "latin-1", "cp1252"):
                try:
                    texto_arch = contenido_bytes.decode(enc)
                    break
                except UnicodeDecodeError:
                    texto_arch = contenido_bytes.decode("utf-8", errors="ignore")
            bloques.append(f"=== ARCHIVO: {archivo.name} ===\n{texto_arch}")
        elif nombre.endswith((".xlsx", ".xls")):
            hojas = pd.read_excel(io.BytesIO(contenido_bytes), sheet_name=None)
            for nombre_hoja, df in hojas.items():
                bloques.append(f"=== EXCEL ({archivo.name} - {nombre_hoja}) ===\n{df.to_csv(index=False)}")
        elif nombre.endswith(".pdf"):
            lector = PdfReader(io.BytesIO(contenido_bytes))
            paginas = [p.extract_text() or "" for p in lector.pages]
            bloques.append(f"=== PDF: {archivo.name} ===\n" + "\n".join(paginas))

    texto_final = "\n\n".join(bloques).strip()
    return "texto", None, texto_final, texto_final


# ---------------------------------------------------------
# 3. CEREBRO ANALISTA CON REINTENTOS Y RESPALDO MULTI-MODELO
# ---------------------------------------------------------
def llamar_gemini_resiliente(api_key: str, prompt: str) -> dict:
    client = genai.Client(api_key=api_key)

    modelos_candidatos = [
        "gemini-3.8-flash",
        "gemini-3.8-flash-lite",
        "gemini-2.5-flash",
    ]
    ultimo_error = None
    descubiertos = False

    idx = 0
    while idx < len(modelos_candidatos):
        modelo = modelos_candidatos[idx]
        idx += 1

        # Hasta 2 intentos por modelo si hay pico temporal 503 / 429
        for intento in range(2):
            try:
                respuesta = client.models.generate_content(
                    model=modelo,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.1,
                    ),
                )
                texto_json = respuesta.text.strip()
                if texto_json.startswith("```"):
                    texto_json = re.sub(r"^```(?:json)?\s*|\s*```$", "", texto_json)
                return json.loads(texto_json)
            except Exception as e:
                ultimo_error = e
                err_str = str(e)
                es_saturado = any(
                    cod in err_str
                    for cod in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "overloaded")
                )
                es_no_encontrado = "404" in err_str or "NOT_FOUND" in err_str

                if es_saturado and intento == 0:
                    time.sleep(2.5)
                    continue

                if es_saturado or es_no_encontrado:
                    if not descubiertos:
                        descubiertos = True
                        try:
                            for m in client.models.list():
                                nombre_m = getattr(m, "name", "").replace("models/", "")
                                if (
                                    "gemini" in nombre_m
                                    and not any(
                                        x in nombre_m
                                        for x in ("audio", "image", "tts", "embedding", "custom", "vision")
                                    )
                                    and nombre_m not in modelos_candidatos
                                ):
                                    modelos_candidatos.append(nombre_m)
                        except Exception:
                            pass
                    break  # Pasar al siguiente modelo candidato

                raise e

    raise RuntimeError(
        f"Los servidores gratuitos de Google tienen alta demanda en este momento. "
        f"Esperá 10 segundos y volvé a apretar el botón. (Detalle: {ultimo_error})"
    )


def analizar_con_ia(
    api_key: str,
    modo: str,
    df_tabular: pd.DataFrame,
    contenido_ia: str,
    contexto_extra: str,
) -> dict:
    if modo == "tabular":
        cols_exactas = list(df_tabular.columns)
        prompt = f"""
Sos un Analista de Datos Senior. El usuario subió una planilla ya estructurada con {len(df_tabular)} filas.
Te paso las columnas exactas, las estadísticas reales calculadas sobre el 100% de los datos y una muestra de 20 filas.

COLUMNAS EXACTAS DISPONIBLES:
{json.dumps(cols_exactas, ensure_ascii=False)}

CONTEXTO ADICIONAL DEL USUARIO:
{contexto_extra if contexto_extra else "Ninguno. Deducí la mejor lógica de negocio posible."}

TAREA:
1. Elegí UNA 'columna_categoria' (debe ser exactamente uno de los nombres en COLUMNAS EXACTAS DISPONIBLES, idealmente texto como Proveedor, Rubro, Cliente, Vendedor, Estado o Sucursal).
2. Elegí UNA 'columna_metrica' (debe ser exactamente uno de los nombres numéricos en COLUMNAS EXACTAS DISPONIBLES, como Importe, Total, Monto, Cantidad).
3. Indicá si 'es_moneda' es true o false.
4. Redactá entre 3 y 5 hallazgos ejecutivos concretos ('resumen_ejecutivo') usando las cifras reales de las estadísticas proporcionadas.

Devolvé ÚNICAMENTE un JSON válido con esta estructura exacta (NO devuelvas las filas porque ya las tenemos en Python):
{{
  "titulo_reporte": "Título ejecutivo corto",
  "tipo_detectado": "Ej: Reporte de Compras / Ventas / Stock",
  "resumen_ejecutivo": [
    "Hallazgo 1 con cifras reales",
    "Hallazgo 2 con cifras reales",
    "Hallazgo 3 con cifras reales"
  ],
  "columna_categoria": "Nombre exacto de columna",
  "columna_metrica": "Nombre exacto de columna numérica",
  "es_moneda": true
}}

DATOS Y ESTADÍSTICAS:
{contenido_ia[:30000]}
"""
        resultado = llamar_gemini_resiliente(api_key, prompt)
        resultado["columnas"] = cols_exactas
        resultado["filas"] = df_tabular.values.tolist()
        return resultado

    # Modo texto libre / PDF
    prompt = f"""
Sos un Analista de Datos Senior. Vas a recibir datos crudos provenientes de texto copiado, PDF o archivos sin estructura tabular fija.
Tu trabajo es estructurarlos para armar un libro de Excel gerencial impecable.

CONTEXTO ADICIONAL DEL USUARIO:
{contexto_extra if contexto_extra else "Ninguno. Deducí la mejor lógica de negocio posible."}

REGLAS ESTRICTAS:
1. Limpiá encabezados basura, totales repetidos y filas vacías.
2. Normalizá números (sin signos $ ni separadores de miles: devolvelos como float o int puros en el JSON).
3. Normalizá fechas al formato YYYY-MM-DD cuando existan.
4. Elegí UNA columna categórica principal ('columna_categoria') y UNA columna numérica principal ('columna_metrica'). Si no hay números en los datos, creá una columna 'Registros' con valor 1 en cada fila.
5. Redactá entre 3 y 5 hallazgos ejecutivos concretos ('resumen_ejecutivo') con cifras reales calculadas sobre los datos.

Devolvé ÚNICAMENTE un JSON válido con esta estructura exacta:
{{
  "titulo_reporte": "Título ejecutivo corto",
  "tipo_detectado": "Ej: Reporte de Ventas / Extracto Bancario / Control de Stock",
  "resumen_ejecutivo": [
    "Hallazgo 1 con datos concretos",
    "Hallazgo 2 con datos concretos",
    "Hallazgo 3 con datos concretos"
  ],
  "columna_categoria": "Nombre exacto de la columna en 'columnas'",
  "columna_metrica": "Nombre exacto de la columna numérica en 'columnas'",
  "es_moneda": true,
  "columnas": ["Col1", "Col2", "Col3"],
  "filas": [
    ["valor1", "valor2", 1500.50]
  ]
}}

DATOS CRUDOS A PROCESAR:
{contenido_ia[:35000]}
"""
    return llamar_gemini_resiliente(api_key, prompt)


# ---------------------------------------------------------
# 4. GENERADOR DE EXCEL CON FÓRMULAS Y GRÁFICOS REALES
# ---------------------------------------------------------
def construir_excel_profesional(analisis: dict, datos_crudos: str) -> bytes:
    wb = openpyxl.Workbook()

    azul_oscuro = PatternFill("solid", fgColor="1B365D")
    azul_medio = PatternFill("solid", fgColor="2C5282")
    gris_suave = PatternFill("solid", fgColor="F4F6F9")
    tarjeta_fill = PatternFill("solid", fgColor="EDF2F7")

    font_titulo = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    font_subtitulo = Font(name="Calibri", size=10, italic=True, color="E2E8F0")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_kpi_label = Font(name="Calibri", size=9, bold=True, color="4A5568")
    font_kpi_val = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_seccion = Font(name="Calibri", size=12, bold=True, color="1B365D")
    font_bold = Font(name="Calibri", size=11, bold=True)

    borde_fino = Border(
        left=Side(style="thin", color="CBD5E0"),
        right=Side(style="thin", color="CBD5E0"),
        top=Side(style="thin", color="CBD5E0"),
        bottom=Side(style="thin", color="CBD5E0"),
    )

    columnas = analisis.get("columnas", [])
    filas = analisis.get("filas", [])
    if not columnas:
        columnas = ["Dato"]
    if not filas:
        filas = [["Sin registros"]]

    col_cat = analisis.get("columna_categoria", columnas[0])
    col_met = analisis.get("columna_metrica", columnas[-1])
    es_moneda = bool(analisis.get("es_moneda", False))
    formato_num = "$#,##0.00" if es_moneda else "#,##0.00"

    if col_cat not in columnas:
        col_cat = columnas[0]
    if col_met not in columnas:
        col_met = columnas[-1]

    idx_cat = columnas.index(col_cat) + 1
    idx_met = columnas.index(col_met) + 1
    letra_cat = get_column_letter(idx_cat)
    letra_met = get_column_letter(idx_met)
    total_filas = max(len(filas), 1)
    fila_fin_datos = total_filas + 1

    ws_resumen = wb.active
    ws_resumen.title = "Resumen Ejecutivo"
    ws_analisis = wb.create_sheet("Análisis Dinámico")
    ws_datos = wb.create_sheet("Base Limpia")
    ws_crudo = wb.create_sheet("Crudo Original")

    # PESTAÑA 3: BASE LIMPIA
    for c_idx, col_name in enumerate(columnas, start=1):
        celda = ws_datos.cell(row=1, column=c_idx, value=col_name)
        celda.font = font_header
        celda.fill = azul_oscuro
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_fino

    for r_idx, fila in enumerate(filas, start=2):
        for c_idx in range(1, len(columnas) + 1):
            val = fila[c_idx - 1] if c_idx - 1 < len(fila) else ""
            if isinstance(val, str) and c_idx == idx_met:
                conv = normalizar_numero(val)
                if isinstance(conv, (int, float)):
                    val = conv
            celda = ws_datos.cell(row=r_idx, column=c_idx, value=val)
            celda.border = borde_fino
            if r_idx % 2 == 0:
                celda.fill = gris_suave
            if isinstance(val, (int, float)):
                celda.number_format = formato_num if c_idx == idx_met else "#,##0.##"

    ws_datos.freeze_panes = "A2"
    ws_datos.auto_filter.ref = f"A1:{get_column_letter(len(columnas))}{fila_fin_datos}"

    # PESTAÑA 2: ANÁLISIS DINÁMICO (CON FÓRMULAS REALES)
    ws_analisis.merge_cells("A1:E1")
    ws_analisis["A1"] = f"ANÁLISIS POR {col_cat.upper()}"
    ws_analisis["A1"].font = font_titulo
    ws_analisis["A1"].fill = azul_oscuro
    ws_analisis["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_analisis.row_dimensions[1].height = 30

    headers_ana = [
        col_cat,
        "Cantidad Registros",
        f"Total {col_met}",
        f"Promedio {col_met}",
        "% Participación",
    ]
    for c_idx, h in enumerate(headers_ana, start=1):
        c = ws_analisis.cell(row=3, column=c_idx, value=h)
        c.font = font_header
        c.fill = azul_medio
        c.alignment = Alignment(horizontal="center")
        c.border = borde_fino

    # Top 25 categorías ordenadas por frecuencia/aparición
    conteo_cats = {}
    for f in filas:
        if idx_cat - 1 < len(f):
            v = str(f[idx_cat - 1]).strip()
            if v:
                conteo_cats[v] = conteo_cats.get(v, 0) + 1
    categorias_unicas = [
        k for k, _ in sorted(conteo_cats.items(), key=lambda x: x[1], reverse=True)
    ][:25]

    fila_inicio_ana = 4
    fila_total_ana = fila_inicio_ana + len(categorias_unicas)

    for i, cat in enumerate(categorias_unicas):
        r = fila_inicio_ana + i
        c1 = ws_analisis.cell(row=r, column=1, value=cat)
        c2 = ws_analisis.cell(
            row=r,
            column=2,
            value=f"=COUNTIF('Base Limpia'!${letra_cat}$2:${letra_cat}${fila_fin_datos}, A{r})",
        )
        c3 = ws_analisis.cell(
            row=r,
            column=3,
            value=f"=SUMIF('Base Limpia'!${letra_cat}$2:${letra_cat}${fila_fin_datos}, A{r}, 'Base Limpia'!${letra_met}$2:${letra_met}${fila_fin_datos})",
        )
        c4 = ws_analisis.cell(row=r, column=4, value=f"=IF(B{r}>0, C{r}/B{r}, 0)")
        c5 = ws_analisis.cell(
            row=r, column=5, value=f"=IF($C${fila_total_ana}>0, C{r}/$C${fila_total_ana}, 0)"
        )

        c2.number_format = "#,##0"
        c3.number_format = formato_num
        c4.number_format = formato_num
        c5.number_format = "0.0%"

        for c in (c1, c2, c3, c4, c5):
            c.border = borde_fino

    if categorias_unicas:
        ws_analisis.cell(row=fila_total_ana, column=1, value="TOTAL TOP CATEGORÍAS").font = font_bold
        c_tot_reg = ws_analisis.cell(
            row=fila_total_ana, column=2, value=f"=SUM(B{fila_inicio_ana}:B{fila_total_ana - 1})"
        )
        c_tot_met = ws_analisis.cell(
            row=fila_total_ana, column=3, value=f"=SUM(C{fila_inicio_ana}:C{fila_total_ana - 1})"
        )
        c_tot_prom = ws_analisis.cell(
            row=fila_total_ana,
            column=4,
            value=f"=IF(B{fila_total_ana}>0, C{fila_total_ana}/B{fila_total_ana}, 0)",
        )
        c_tot_pct = ws_analisis.cell(
            row=fila_total_ana, column=5, value=f"=SUM(E{fila_inicio_ana}:E{fila_total_ana - 1})"
        )

        c_tot_reg.number_format = "#,##0"
        c_tot_met.number_format = formato_num
        c_tot_prom.number_format = formato_num
        c_tot_pct.number_format = "0.0%"

        for col_i in range(1, 6):
            cell_t = ws_analisis.cell(row=fila_total_ana, column=col_i)
            cell_t.font = font_bold
            cell_t.fill = tarjeta_fill
            cell_t.border = borde_fino

        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = f"{col_met} por {col_cat}"
        chart.y_axis.title = col_met
        chart.x_axis.title = col_cat
        data_ref = Reference(ws_analisis, min_col=3, min_row=3, max_row=fila_total_ana - 1)
        cats_ref = Reference(ws_analisis, min_col=1, min_row=4, max_row=fila_total_ana - 1)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.legend = None
        chart.width = 18
        chart.height = 11
        ws_analisis.add_chart(chart, "G3")

    # PESTAÑA 1: RESUMEN EJECUTIVO
    ws_resumen.merge_cells("A1:F1")
    ws_resumen["A1"] = analisis.get("titulo_reporte", "Reporte Analítico").upper()
    ws_resumen["A1"].font = font_titulo
    ws_resumen["A1"].fill = azul_oscuro
    ws_resumen["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_resumen.row_dimensions[1].height = 34

    ws_resumen.merge_cells("A2:F2")
    ws_resumen["A2"] = (
        f"Tipo detectado: {analisis.get('tipo_detectado', 'General')}  |  "
        f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    ws_resumen["A2"].font = font_subtitulo
    ws_resumen["A2"].fill = azul_medio
    ws_resumen["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_resumen.row_dimensions[2].height = 22

    kpis = [
        ("A", "B", f"TOTAL {col_met.upper()}", f"=SUM('Base Limpia'!{letra_met}2:{letra_met}{fila_fin_datos})", formato_num),
        ("C", "D", "REGISTROS PROCESADOS", f"=COUNTA('Base Limpia'!A2:A{fila_fin_datos})", "#,##0"),
        ("E", "F", f"PROMEDIO {col_met.upper()}", f"=AVERAGE('Base Limpia'!{letra_met}2:{letra_met}{fila_fin_datos})", formato_num),
    ]
    ws_resumen.row_dimensions[4].height = 20
    ws_resumen.row_dimensions[5].height = 30

    for c_ini, c_fin, etiqueta, formula, num_fmt in kpis:
        ws_resumen.merge_cells(f"{c_ini}4:{c_fin}4")
        ws_resumen.merge_cells(f"{c_ini}5:{c_fin}5")
        lbl_cell = ws_resumen[f"{c_ini}4"]
        val_cell = ws_resumen[f"{c_ini}5"]
        lbl_cell.value = etiqueta
        lbl_cell.font = font_kpi_label
        lbl_cell.alignment = Alignment(horizontal="center", vertical="center")
        val_cell.value = formula
        val_cell.font = font_kpi_val
        val_cell.number_format = num_fmt
        val_cell.alignment = Alignment(horizontal="center", vertical="center")
        for col_let in (c_ini, c_fin):
            ws_resumen[f"{col_let}4"].fill = tarjeta_fill
            ws_resumen[f"{col_let}5"].fill = tarjeta_fill
            ws_resumen[f"{col_let}4"].border = borde_fino
            ws_resumen[f"{col_let}5"].border = borde_fino

    ws_resumen["A7"] = "CONCLUSIONES CLAVE DEL ANALISTA"
    ws_resumen["A7"].font = font_seccion

    insights = analisis.get("resumen_ejecutivo", [])
    for idx, texto_ins in enumerate(insights, start=8):
        ws_resumen.merge_cells(f"A{idx}:F{idx}")
        celda_ins = ws_resumen[f"A{idx}"]
        celda_ins.value = f"•  {texto_ins}"
        celda_ins.alignment = Alignment(wrap_text=True, vertical="center")
        ws_resumen.row_dimensions[idx].height = 32

    # PESTAÑA 4: CRUDO ORIGINAL
    ws_crudo["A1"] = "RESPALDO DE DATOS ORIGINALES INGRESADOS"
    ws_crudo["A1"].font = font_bold
    for idx_linea, linea in enumerate(datos_crudos.splitlines()[:1000], start=3):
        ws_crudo.cell(row=idx_linea, column=1, value=linea)

    for hoja in (ws_resumen, ws_analisis, ws_datos):
        for col in hoja.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value and not str(cell.value).startswith("="):
                    max_len = max(max_len, len(str(cell.value)))
            hoja.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 42)

    salida = io.BytesIO()
    wb.save(salida)
    return salida.getvalue()


# ---------------------------------------------------------
# 5. INTERFAZ PRINCIPAL DE TRABAJO
# ---------------------------------------------------------
def main():
    if not verificar_login():
        return

    with st.sidebar:
        st.success(f"👤 Conectado como: **{st.session_state.usuario_actual}**")
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state.autenticado = False
            st.session_state.usuario_actual = None
            st.rerun()

        st.divider()
        st.subheader("🔑 Tu clave de IA personal")
        api_key = st.text_input(
            "Pegá tu API Key de Google Gemini",
            type="password",
            help="Se guarda solo en tu navegador durante esta sesión. No consume la cuenta de la empresa.",
        )
        with st.expander("¿No tenés clave? Sacala gratis en 1 min"):
            st.markdown(
                "1. Entrá a [aistudio.google.com](https://aistudio.google.com/apikey)\n"
                "2. Iniciá sesión con tu Gmail personal.\n"
                "3. Tocá **Create API Key**, copiala y pegala acá arriba."
            )

    st.title("📊 Analista Excel Automático")
    st.write(
        "Pegá texto suelto (WhatsApp, mails, tablas) o subí cualquier exportación "
        "(CSV, Excel, PDF, TXT, JSON) para obtener un Excel estructurado con KPIs, fórmulas y gráficos."
    )

    col1, col2 = st.columns([1.2, 1])
    with col1:
        texto_usuario = st.text_area(
            "1. Pegá texto o datos crudos acá (opcional si subís archivo)",
            height=200,
            placeholder="Ej: 12/11 Juan Perez - 3 cajas vino $45.000...",
        )
    with col2:
        archivo_subido = st.file_uploader(
            "2. O adjuntá un archivo exportado",
            type=["csv", "xlsx", "xls", "pdf", "txt", "json"],
        )
        contexto = st.text_input(
            "3. Indicaciones extra para el analista (opcional)",
            placeholder="Ej: Agrupá por sucursal y calculá comisiones del 5%",
        )

    if st.button("🚀 Analizar y Generar Excel", type="primary", use_container_width=True):
        if not api_key:
            st.warning("⚠️ Pegá tu API Key personal en el panel de la izquierda antes de procesar.")
            return

        if not (texto_usuario and texto_usuario.strip()) and archivo_subido is None:
            st.warning("⚠️ Pegá algún texto o subí un archivo primero.")
            return

        with st.spinner("El analista IA está limpiando los datos y armando las fórmulas..."):
            try:
                modo, df_tabular, contenido_ia, respaldo_crudo = preparar_entrada(
                    texto_usuario, archivo_subido
                )
                resultado_ia = analizar_con_ia(
                    api_key, modo, df_tabular, contenido_ia, contexto
                )
                excel_bytes = construir_excel_profesional(resultado_ia, respaldo_crudo)

                st.success(
                    f"✅ ¡Excel analítico listo! ({len(resultado_ia.get('filas', []))} registros procesados)"
                )
                st.subheader(resultado_ia.get("titulo_reporte", "Resumen del Análisis"))

                for insight in resultado_ia.get("resumen_ejecutivo", []):
                    st.markdown(f"- {insight}")

                nombre_salida = f"Analisis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                st.download_button(
                    label="📥 Descargar Excel Elaborado (.xlsx)",
                    data=excel_bytes,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Ocurrió un error al procesar: {e}")


if __name__ == "__main__":
    main()
