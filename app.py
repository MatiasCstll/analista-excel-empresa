import io
import json
import re
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
# 2. EXTRACCIÓN UNIVERSAL DE ENTRADA (TEXTO O ARCHIVOS)
# ---------------------------------------------------------
def extraer_texto_entrada(texto_pegado: str, archivo) -> str:
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
            excel_io = io.BytesIO(contenido_bytes)
            hojas = pd.read_excel(excel_io, sheet_name=None)
            for nombre_hoja, df in hojas.items():
                csv_hoja = df.to_csv(index=False)
                bloques.append(f"=== EXCEL ({archivo.name} - Hoja: {nombre_hoja}) ===\n{csv_hoja}")

        elif nombre.endswith(".pdf"):
            lector = PdfReader(io.BytesIO(contenido_bytes))
            paginas = [p.extract_text() or "" for p in lector.pages]
            bloques.append(f"=== PDF: {archivo.name} ===\n" + "\n".join(paginas))

    return "\n\n".join(bloques).strip()


# ---------------------------------------------------------
# 3. CEREBRO ANALISTA (USA LA API KEY PROPIA DEL EMPLEADO)
# ---------------------------------------------------------
def analizar_con_ia(api_key: str, datos_crudos: str, contexto_extra: str) -> dict:
    client = genai.Client(api_key=api_key)

    prompt = f"""
Sos un Analista de Datos Senior. Vas a recibir datos crudos provenientes de cualquier sistema, planilla, PDF o texto copiado.
Tu trabajo es estructurarlos para armar un libro de Excel gerencial impecable.

CONTEXTO ADICIONAL DEL USUARIO:
{contexto_extra if contexto_extra else "Ninguno. Deducí la mejor lógica de negocio posible."}

REGLAS ESTRICTAS:
1. Limpiá encabezados basura, totales repetidos y filas vacías.
2. Normalizá números (sin signos $ ni separadores de miles: devolvelos como float o int puros en el JSON).
3. Normalizá fechas al formato YYYY-MM-DD cuando existan.
4. Elegí UNA columna categórica principal ('columna_categoria', ej: Vendedor, Cliente, Rubro, Estado, Sucursal) para agrupar en tablas dinámicas.
5. Elegí UNA columna numérica principal ('columna_metrica', ej: Importe, Total, Cantidad, Horas) para sumar y promediar. Si no hay números en los datos, creá una columna 'Registros' con valor 1 en cada fila.
6. Redactá entre 3 y 5 hallazgos ejecutivos concretos ('resumen_ejecutivo') con cifras reales calculadas sobre los datos.

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
{datos_crudos[:45000]}
"""

    modelos = ["gemini-3.8-flash"]
    ultimo_error = None
    buscados = False

    idx = 0
    while idx < len(modelos):
        modelo = modelos[idx]
        idx += 1
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
            if "404" in err_str or "NOT_FOUND" in err_str:
                if not buscados:
                    buscados = True
                    try:
                        for m in client.models.list():
                            nombre_m = getattr(m, "name", "").replace("models/", "")
                            if (
                                "flash" in nombre_m
                                and not any(
                                    x in nombre_m
                                    for x in ("audio", "image", "tts", "embedding", "custom")
                                )
                                and nombre_m not in modelos
                            ):
                                modelos.append(nombre_m)
                    except Exception:
                        pass
                continue
            raise e

    raise ultimo_error


# ---------------------------------------------------------
# 4. GENERADOR DE EXCEL CON FÓRMULAS Y GRÁFICOS REALES
# ---------------------------------------------------------
def construir_excel_profesional(analisis: dict, datos_crudos: str) -> bytes:
    wb = openpyxl.Workbook()

    # Estilos corporativos
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
    col_cat = analisis.get("columna_categoria", columnas[0] if columnas else "Categoría")
    col_met = analisis.get("columna_metrica", columnas[-1] if columnas else "Valor")
    es_moneda = bool(analisis.get("es_moneda", False))
    formato_num = "$#,##0.00" if es_moneda else "#,##0.00"

    if col_cat not in columnas and columnas:
        col_cat = columnas[0]
    if col_met not in columnas and columnas:
        col_met = columnas[-1]

    idx_cat = columnas.index(col_cat) + 1
    idx_met = columnas.index(col_met) + 1
    letra_cat = get_column_letter(idx_cat)
    letra_met = get_column_letter(idx_met)
    total_filas = max(len(filas), 1)
    fila_fin_datos = total_filas + 1

    # Crear las 4 pestañas
    ws_resumen = wb.active
    ws_resumen.title = "Resumen Ejecutivo"
    ws_analisis = wb.create_sheet("Análisis Dinámico")
    ws_datos = wb.create_sheet("Base Limpia")
    ws_crudo = wb.create_sheet("Crudo Original")

    # -----------------------------------------------------
    # PESTAÑA 3: BASE LIMPIA
    # -----------------------------------------------------
    for c_idx, col_name in enumerate(columnas, start=1):
        celda = ws_datos.cell(row=1, column=c_idx, value=col_name)
        celda.font = font_header
        celda.fill = azul_oscuro
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_fino

    for r_idx, fila in enumerate(filas, start=2):
        for c_idx in range(1, len(columnas) + 1):
            val = fila[c_idx - 1] if c_idx - 1 < len(fila) else ""
            if isinstance(val, str):
                limpio = re.sub(r"[^\d.,-]", "", val)
                if c_idx == idx_met and limpio:
                    try:
                        val = float(limpio.replace(",", "."))
                    except ValueError:
                        pass
            celda = ws_datos.cell(row=r_idx, column=c_idx, value=val)
            celda.border = borde_fino
            if r_idx % 2 == 0:
                celda.fill = gris_suave
            if isinstance(val, (int, float)):
                celda.number_format = formato_num if c_idx == idx_met else "#,##0.##"

    ws_datos.freeze_panes = "A2"
    ws_datos.auto_filter.ref = f"A1:{get_column_letter(len(columnas))}{fila_fin_datos}"

    # -----------------------------------------------------
    # PESTAÑA 2: ANÁLISIS DINÁMICO (CON FÓRMULAS REALES)
    # -----------------------------------------------------
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

    categorias_unicas = []
    for f in filas:
        if idx_cat - 1 < len(f):
            v = str(f[idx_cat - 1]).strip()
            if v and v not in categorias_unicas:
                categorias_unicas.append(v)
    categorias_unicas = categorias_unicas[:25]

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

    # Fila de Totales
    if categorias_unicas:
        ws_analisis.cell(row=fila_total_ana, column=1, value="TOTAL GENERAL").font = font_bold
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

        # Gráfico nativo de Excel
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

    # -----------------------------------------------------
    # PESTAÑA 1: RESUMEN EJECUTIVO
    # -----------------------------------------------------
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

    # Tarjetas KPI con fórmulas dinámicas
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

    # -----------------------------------------------------
    # PESTAÑA 4: CRUDO ORIGINAL
    # -----------------------------------------------------
    ws_crudo["A1"] = "RESPALDO DE DATOS ORIGINALES INGRESADOS"
    ws_crudo["A1"].font = font_bold
    for idx_linea, linea in enumerate(datos_crudos.splitlines()[:2000], start=3):
        ws_crudo.cell(row=idx_linea, column=1, value=linea)

    # Autoajuste de columnas
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

        datos_crudos = extraer_texto_entrada(texto_usuario, archivo_subido)
        if not datos_crudos:
            st.warning("⚠️ Pegá algún texto o subí un archivo primero.")
            return

        with st.spinner("El analista IA está limpiando los datos y armando las fórmulas..."):
            try:
                resultado_ia = analizar_con_ia(api_key, datos_crudos, contexto)
                excel_bytes = construir_excel_profesional(resultado_ia, datos_crudos)

                st.success("✅ ¡Excel analítico listo!")
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
