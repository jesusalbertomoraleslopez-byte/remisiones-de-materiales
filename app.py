import streamlit as st
import pandas as pd
import datetime
import io
import requests
import base64
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# 1. CONFIGURACIÓN E INTERFAZ BASE RESPONSIVA
st.set_page_config(page_title="Remisiones de Materiales", layout="wide", page_icon="📦")

# Renderizado de Banner Corporativo Adaptable
try:
    banner_img = Image.open("REMISIONES APP.png")
    st.image(banner_img, use_container_width=True)
except FileNotFoundError:
    st.warning("⚠️ Cargando interfaz gráfica del banner superior corporativo...")
st.write("")

# =============================================================================
# 2. MOTOR DE PERSISTENCIA: LECTURA DE MAESTROS DESDE TU REPOSITORIO DE GITHUB
# =============================================================================
REPO_OWNER = "jesusalbertomoraleslopez-byte"
REPO_NAME = "remisiones-de-materiales"
BRANCH = "main"

def cargar_excel_desde_github(file_name):
    """Descarga el archivo Excel desde GitHub en crudo si existe, de lo contrario devuelve None."""
    try:
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{file_name}"
        res = requests.get(raw_url)
        if res.status_code == 200:
            return pd.read_excel(io.BytesIO(res.content))
    except Exception:
        pass
    return None
# --- ENLACE DINÁMICO DE HOJAS ELECTRÓNICAS AL ENTORNO DE TRABAJO ---
if "BD_Articulos" not in st.session_state:
    st.session_state.BD_Articulos = pd.DataFrame([
        {"SKU": "12-B-9016-01", "Nombre": "Lámina Galvanizada Sigrama", "Calibre_Espesor": "Calibre 22", "Dimensiones_Pieza": "3x10 ft", "Acabado_Superficial": "Zintro"},
        {"SKU": "SKU-002", "Nombre": "Placa de Acero Comercial", "Calibre_Espesor": "1/4 pulgada", "Dimensiones_Pieza": "4x8 ft", "Acabado_Superficial": "Negro"}
    ])

if "BD_Tarimas" not in st.session_state:
    df_git = cargar_excel_desde_github("BD_Tarimas.xlsx")
    if df_git is not None:
        st.session_state.BD_Tarimas = df_git
    else:
        st.session_state.BD_Tarimas = pd.DataFrame(columns=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus", "Es_Nueva"])

if "BD_Detalle_Tarimas" not in st.session_state:
    df_git = cargar_excel_desde_github("BD_Detalle_Tarimas.xlsx")
    if df_git is not None:
        st.session_state.BD_Detalle_Tarimas = df_git
    else:
        st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Proyecto", "Parcialidad", "Descripcion", "Cantidad"])
if "BD_Datos_Generales_Remision" not in st.session_state:
    df_git = cargar_excel_desde_github("BD_Datos_Generales_Remision.xlsx")
    if df_git is not None:
        st.session_state.BD_Datos_Generales_Remision = df_git
    else:
        st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=["ID_Remision", "Folio_Remision", "Fecha_Hora_Salida", "Nombre_Emisor", "Direccion_Emisor", "Nombre_Receptor", "Direccion_Receptor", "Tarimas_Asociadas"])

if "BD_Lideres" not in st.session_state:
    df_git = cargar_excel_desde_github("BD_Lideres.xlsx")
    if df_git is not None:
        st.session_state.BD_Lideres = df_git
    else:
        st.session_state.BD_Lideres = pd.DataFrame([{"ID_Lider": "LID-01", "Nombre_Lider": "Jesus Morales", "Area": "Metales", "Estatus": "Activo"}])

# CONECTOR MAESTRO DE INYECCIÓN DE RESPALDOS AUTOMÁTICOS HACIA GITHUB
def subir_excel_a_github(file_name, dataframe_to_save):
    """Sincroniza y escribe directamente el DataFrame en el repositorio mediante la API de GitHub."""
    try:
        GITHUB_TOKEN = st.secrets["github_token"]
        url = f"https://github.com{REPO_OWNER}/{REPO_NAME}/contents/{file_name}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
        
        buffer_git = io.BytesIO()
        with pd.ExcelWriter(buffer_git, engine='openpyxl') as writer:
            dataframe_to_save.to_excel(writer, index=False, sheet_name='Datos_Sistema')
            
        base64_content = base64.b64encode(buffer_git.getvalue()).decode("utf-8")
        res_get = requests.get(url, headers=headers)
        sha = res_get.json().get("sha") if res_get.status_code == 200 else None
        
        payload = {"message": f"Sincronizacion App: {file_name}", "content": base64_content, "branch": BRANCH}
        if sha: payload["sha"] = sha
        
        res_put = requests.put(url, json=payload, headers=headers)
        return res_put.status_code in [200, 201]
    except Exception:
        return False

# =============================================================================
# CAPA DE CONTROL DOCUMENTAL Y DISEÑO IMPRESO CORPORATIVO (INDUSTRIA SIGRAMA)
# =============================================================================
def draw_sigrama_decorations(canvas, doc):
    """Dibuja los elementos de marca institucionales y códigos de sistema fijos."""
    canvas.saveState()
    # Franja superior roja Sigrama
    canvas.setFillColor(colors.HexColor("#D32F2F"))
    canvas.rect(36, 745, 540, 4, fill=1, stroke=0)
    
    # Marcador de Control de Calidad Superior Izquierdo (FO-MET-10)
    canvas.setFont("Helvetica-Bold", 11)
    canvas.setFillColor(colors.HexColor("#D32F2F"))
    canvas.drawString(36, 765, "FO-MET-10")
    
    # Metadatos de Revisión de Control
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.black)
    canvas.drawString(36, 753, "Revisión 01")
    canvas.drawString(36, 741, "04 de octubre 2018")
    
    # Título Central del Formato Oficial
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawCentredString(285, 755, "EMBARQUE-RECEPCIÓN DE MERCANCÍA")
    
    # Pie de Página Legal y Control del SGC (FO-SGC-02)
    canvas.setStrokeColor(colors.HexColor("#D32F2F"))
    canvas.setLineWidth(1)
    canvas.line(36, 45, 36, 25)
    canvas.setFont("Helvetica-Bold", 7)
    canvas.drawString(42, 37, "FO-SGC-02")
    
    canvas.setFont("Helvetica", 6)
    canvas.setFillColor(colors.HexColor("#424242"))
    texto_legal = "PROHIBIDA LA REPRODUCCIÓN TOTAL O PARCIAL, POR CUALQUIER MEDIO O PROCEDIMIENTO, SIN AUTORIZACIÓN DE INDUSTRIA SIGRAMA S.A. DE C.V."
    canvas.drawString(95, 37, texto_legal)
    canvas.restoreState()

def generar_pdf_tarima(id_tarima):
    """Genera la carátula de identificación y desglose interno para un bulto/tarima específico."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    # Extracción segura de metadatos de almacenamiento relacional
    tarima_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == id_tarima].iloc[0]
    detalles = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == id_tarima]
    
    # Hoja 1: Indicador gigante para identificación rápida en patio/almacén
    style_g = ParagraphStyle('G_Tar', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1)
    story.append(Spacer(1, 1.8 * inch))
    story.append(Paragraph(f"TARIMA<br/><br/><b>#{id_tarima}</b>", style_g))
    story.append(PageBreak())
    
    # Hoja 2: Ficha técnica de control interno y cantidades
    style_n = styles['Normal']
    style_ng = ParagraphStyle('NG_Tar', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)
    story.append(Paragraph(f"<b>Detalle Interno - Tarima #{id_tarima}</b>", styles['Heading2']))
    story.append(Paragraph(f"<b>Operador / Líder:</b> {tarima_info['Creado_Por']} | <b>Fecha:</b> {tarima_info['Fecha_Creacion']}", style_n))
    story.append(Spacer(1, 0.3 * inch))
    
    for _, item in detalles.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
        story.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", style_n))
        story.append(Spacer(1, 0.4 * inch))
        story.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", style_ng))
        
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer


def generar_pdf_remision_general(datos_remision, df_detalles_remision):
    """Construye el documento oficial de despacho con panel logístico, cuadrícula de materiales y firmas."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    style_blanco_bold = ParagraphStyle('BB_R', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_normal_bold = ParagraphStyle('NB_R', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=9)
    style_normal_text = ParagraphStyle('NT_R', parent=styles['Normal'], fontSize=9)
    
    story.append(Spacer(1, 0.1 * inch))
    t_header_embarque = Table([[Paragraph("LISTADO DE EMBARQUE", style_blanco_bold)]], colWidths=[7.5 * inch])
    t_header_embarque.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header_embarque)
    
    # Panel logístico estructurado
    datos_panel = [
        [Paragraph("FOLIO:", style_normal_bold), Paragraph(str(datos_remision['Folio_Remision']), style_normal_text), Paragraph("FECHA:", style_normal_bold), Paragraph(str(datos_remision['Fecha_Hora_Salida']), style_normal_text)],
        [Paragraph("LÍDER:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Emisor']), style_normal_text), Paragraph("ALMACÉN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Emisor']), style_normal_text)],
        [Paragraph("DESTINO / CLIENTE:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Receptor']), style_normal_text), "", ""],
        [Paragraph("DIRECCIÓN DESTINO:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Receptor']), style_normal_text), "", ""]
    ]
    t_panel = Table(datos_panel, colWidths=[1.5 * inch, 2.25 * inch, 1.5 * inch, 2.25 * inch])
    t_panel.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")), ('SPAN', (1,2), (3,2)), ('SPAN', (1,3), (3,3)), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(t_panel)
    story.append(Spacer(1, 0.2 * inch))
    
    # Cuadrícula formal de listado de materiales
    tabla_materiales = [[Paragraph("CANTIDAD", style_blanco_bold), Paragraph("UNIDAD", style_blanco_bold), Paragraph("CLAVE / MODELO", style_blanco_bold), Paragraph("DESCRIPCIÓN", style_blanco_bold), Paragraph("OBSERVACIONES", style_blanco_bold)]]
    for _, row in df_detalles_remision.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == row['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Material de Embarque"
        desc_proy = row['Descripcion'] if 'Descripcion' in row else "Metales"
        tabla_materiales.append([Paragraph(str(int(row['Cantidad'])), style_normal_text), Paragraph("Piezas", style_normal_text), Paragraph(str(row['SKU']), style_normal_text), Paragraph(nom_art, style_normal_text), Paragraph(f"Tarima: {row['ID_Tarima']} | {desc_proy}", style_normal_text)])
    
    # Rellenar renglones vacíos para simular el formato impreso
    for _ in range(max(1, 8 - len(df_detalles_remision))):
        tabla_materiales.append([Paragraph("", style_normal_text)] * 5)
        
    t_mat = Table(tabla_materiales, colWidths=[1.0 * inch, 0.8 * inch, 1.5 * inch, 2.4 * inch, 1.8 * inch])
    t_mat.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(t_mat)
    story.append(Spacer(1, 0.4 * inch))
    
    # Panel inferior de firmas de validación
    t_firmas = Table([[Paragraph("ENTREGA (LÍDER ALMACÉN)", style_normal_bold), Paragraph("RECIBE (CHOFER / CLIENTE)", style_normal_bold)]], colWidths=[3.75 * inch, 3.75 * inch])
    t_firmas.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('LINEBELOW', (0,0), (0,0), 1, colors.black), ('LINEBELOW', (1,0), (1,0), 1, colors.black), ('TOPPADDING', (0,0), (-1,-1), 25)]))
    story.append(t_firmas)
    
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer

def generar_pdf_anexo_tarimas(lista_tarimas_id, df_detalles_remision):
    """Genera las hojas de desglose técnico complementario para el operador receptor sin empalmes de texto."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    style_b = ParagraphStyle('ABB_A', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_t = ParagraphStyle('ANT_A', parent=styles['Normal'], fontSize=9)
    # Corrección definitiva de interlineado (leading) para separar títulos grandes
    style_caratula = ParagraphStyle('C_Fix', parent=styles['Heading1'], fontSize=42, leading=46, alignment=1)
    
    for t_id in lista_tarimas_id:
        # Hoja 1: Carátula divisoria limpia del anexo
        story.append(Spacer(1, 1.8 * inch))
        story.append(Paragraph(f"ANEXO: TARIMA<br/><b>{t_id}</b>", style_caratula))
        story.append(Spacer(1, 0.4 * inch))
        
        t_bar = Table([["|||||||||||||||||||||||||||||||"], [f"*{t_id}*"]], colWidths=[7.5 * inch])
        t_bar.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('TEXTCOLOR', (0,0), (-1,1), colors.darkgray)]))
        story.append(t_bar)
        story.append(PageBreak())
        
        # Hoja 2: Cuadrícula detallada con sombreado de contraste gris
        story.append(Paragraph(f"<b>DETALLE ESPECÍFICO - TARIMA {t_id}</b>", styles['Heading2']))
        story.append(Spacer(1, 0.1 * inch))
        
        sub_det = df_detalles_remision[df_detalles_remision['ID_Tarima'] == t_id]
        tabla_anexo = [[Paragraph("PO ASOCIADA", style_b), Paragraph("SKU / PRODUCTO", style_b), Paragraph("CANTIDAD", style_b)]]
        for _, s_row in sub_det.iterrows():
            tabla_anexo.append([Paragraph(str(s_row['PO']), style_t), Paragraph(str(s_row['SKU']), style_t), Paragraph(str(int(s_row['Cantidad'])), style_t)])
            
        t_det = Table(tabla_anexo, colWidths=[2.5 * inch, 2.5 * inch, 2.5 * inch])
        t_det.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F5F5F5")), ('GRID', (0,0), (-1,-1), 1, colors.white), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        story.append(t_det)
        story.append(PageBreak())
        
    if story: story.pop()
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer

# =============================================================================
# 8. CAPA DE CONTROL DE ACCESOS Y ENRUTAMIENTO DE NAVEGACIÓN
# =============================================================================
st.sidebar.title("🔐 Control de Acceso")
admin_pass_input = st.sidebar.text_input("Contraseña Administrador:", type="password")

def es_admin():
    try:
        return admin_pass_input == st.secrets["admin_password"]
    except KeyError:
        st.sidebar.error("Error: 'admin_password' no configurado en Secrets.")
        return False

is_admin = es_admin()
if is_admin:
    st.sidebar.success("Modo Administrador Activo")
else:
    st.sidebar.warning("Modo Consulta Activo")

# --- VALIDACIÓN DE SUPERUSUARIO PARA MANTENIMIENTO ---
st.sidebar.write("---")
st.sidebar.title("🛠️ Área de Soporte")
super_pass_input = st.sidebar.text_input("Contraseña de Soporte / IT:", type="password")
is_super = (super_pass_input == "SigramaMetales2025")

if is_super:
    st.sidebar.success("⚡ Modo Superusuario Activo")

st.sidebar.title("🧭 Navegación")
lista_modulos = ["📊 Dashboard e Históricos", "🔍 Centro de Consultas", "📦 Módulo Tarimas", "🚚 Módulo Remisiones"]
if is_super:
    lista_modulos.append("⚙️ Mantenimiento y Catálogos")

opcion_menu = st.sidebar.radio("Seleccione un Módulo:", lista_modulos)
# =============================================================================
# 9. INTERFAZ DE USUARIO: VISUALIZACIÓN DE PANELES OPERATIVOS (DASHBOARD)
# =============================================================================
if opcion_menu == "📊 Dashboard e Históricos":
    st.title("📊 Dashboard Planta Metales Inventario Producto")
    
    # Filtros Maestros Globales en la parte superior del Dashboard
    st.subheader("🔍 Filtros de Control Global")
    df_maestro_dash = st.session_state.BD_Detalle_Tarimas.copy()
    
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        opciones_global_proy = ["Todos"] + df_maestro_dash['Proyecto'].dropna().unique().tolist() if not df_maestro_dash.empty and 'Proyecto' in df_maestro_dash.columns else ["Todos"]
        proy_global_sel = st.selectbox("Filtrar por Proyecto Interno:", opciones_global_proy, key="dash_global_proy_select_unique")
        
    with col_g2:
        if proy_global_sel != "Todos":
            df_maestro_dash = df_maestro_dash[df_maestro_dash['Proyecto'] == proy_global_sel]
        opciones_global_desc = ["Todas"] + df_maestro_dash['Descripcion'].dropna().unique().tolist() if not df_maestro_dash.empty and 'Descripcion' in df_maestro_dash.columns else ["Todas"]
        desc_global_sel = st.selectbox("Filtrar por Descripción de Proyecto Planta Rio:", opciones_global_desc, key="dash_global_desc_select_unique")

    st.write("---")
    
    # Lógica de segmentación cruzada de datos según los filtros superiores
    df_detalles_filtrados = st.session_state.BD_Detalle_Tarimas.copy()
    df_tarimas_filtradas = st.session_state.BD_Tarimas.copy()
    
    if proy_global_sel != "Todos":
        df_detalles_filtrados = df_detalles_filtrados[df_detalles_filtrados['Proyecto'] == proy_global_sel]
        tarimas_validas_f = df_detalles_filtrados['ID_Tarima'].unique()
        df_tarimas_filtradas = df_tarimas_filtradas[df_tarimas_filtradas['ID_Tarima'].isin(tarimas_validas_f)]
        
    if desc_global_sel != "Todas":
        df_detalles_filtrados = df_detalles_filtrados[df_detalles_filtrados['Descripcion'] == desc_global_sel]
        tarimas_validas_f = df_detalles_filtrados['ID_Tarima'].unique()
        df_tarimas_filtradas = df_tarimas_filtradas[df_tarimas_filtradas['ID_Tarima'].isin(tarimas_validas_f)]
    # Recálculo instantáneo de las tarjetas KPI del Dashboard basado en los filtros superiores
    t_tar = len(df_tarimas_filtradas)
    disp = len(df_tarimas_filtradas[df_tarimas_filtradas['Estatus'] == 'Disponible'])
    rem = len(df_tarimas_filtradas[df_tarimas_filtradas['Estatus'] == 'Remesada'])
    
    if not df_detalles_filtrados.empty:
        total_piezas = int(df_detalles_filtrados['Cantidad'].sum())
        id_disp = df_tarimas_filtradas[df_tarimas_filtradas['Estatus'] == 'Disponible']['ID_Tarima']
        piezas_disp = int(df_detalles_filtrados[df_detalles_filtrados['ID_Tarima'].isin(id_disp)]['Cantidad'].sum())
    else:
        total_piezas = piezas_disp = 0

    # Despliegue responsivo de las tarjetas grandes de control
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📦 Total Tarimas", f"{t_tar} Reg.")
    m2.metric("🟢 Tarimas Disponibles", disp)
    m3.metric("🚚 Tarimas Remesadas", rem)
    m4.metric("🔢 Total Piezas en Almacén", f"{total_piezas:,} PZS")
    st.write("---")
    st.subheader("📋 Resumen de Cumplimiento y Parcialidades")
    
    if not st.session_state.BD_Detalle_Tarimas.empty:
        df_completo = pd.merge(st.session_state.BD_Detalle_Tarimas, st.session_state.BD_Tarimas[['ID_Tarima', 'Estatus']], on='ID_Tarima', how='left')
        
        # Sincronizar la tabla del resumen analítico inferior con los filtros globales de arriba
        if proy_global_sel != "Todos": df_completo = df_completo[df_completo['Proyecto'] == proy_global_sel]
        if desc_global_sel != "Todas": df_completo = df_completo[df_completo['Descripcion'] == desc_global_sel]
            
        if not df_completo.empty and all(col in df_completo.columns for col in ['Proyecto', 'Parcialidad', 'Descripcion']):
            resumen_avanzado = df_completo.groupby(['PO', 'Proyecto', 'Parcialidad', 'Descripcion']).agg(
                Cant_Tarimas=('ID_Tarima', 'nunique'), Total_Piezas=('Cantidad', 'sum'),
                Piezas_Disponibles=('Cantidad', lambda x: x[df_completo.loc[x.index, 'Estatus'] == 'Disponible'].sum()),
                Piezas_Remesadas=('Cantidad', lambda x: x[df_completo.loc[x.index, 'Estatus'] == 'Remesada'].sum())
            ).reset_index()
            
            resumen_avanzado['% Avance Salida'] = (resumen_avanzado['Piezas_Remesadas'] / resumen_avanzado['Total_Piezas'] * 100).round(1)
            resumen_avanzado['% Avance Salida'] = resumen_avanzado['% Avance Salida'].apply(lambda x: f"{x}%")
            
            # Renombrar columnas con el encabezado oficial del Proyecto Planta Rio solicitado
            resumen_avanzado.columns = ["Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", "Descripción de Proyecto Planta Rio", "Cant. Tarimas", "Total Piezas", "Piezas Disponibles", "Piezas Remesadas", "% Avance Salida"]
            st.dataframe(resumen_avanzado, use_container_width=True, hide_index=True)
        else:
            st.info("No existen parcialidades registradas para el filtro seleccionado.")
    else:
        st.info("No hay datos cargados para segmentar por proyectos.")
        
    st.write("---")
    st.subheader("📋 Estado Detallado del Inventario Entarimado (Maestro)")
    if not st.session_state.BD_Tarimas.empty:
        vista_dash = st.session_state.BD_Tarimas.drop(columns=["Es_Nueva"], errors="ignore")
        st.dataframe(vista_dash, use_container_width=True)
    else:
        st.info("No hay tarimas registradas en el inventario.")
# =============================================================================
# 10. REPORTE GRANULAR E INVENTARIO DETALLADO POR PIEZA (CENTRO DE CONSULTAS)
# =============================================================================
elif opcion_menu == "🔍 Centro de Consultas":
    st.title("🔍 Reporte Detallado de Inventario por Pieza")
    st.markdown("Utilice los selectores inferiores para filtrar de forma granular el inventario consolidado:")
    
    if not st.session_state.BD_Detalle_Tarimas.empty:
        # Cruce relacional para anexar estatus de envío desde el catálogo de tarimas
        df_maestro_piezas = pd.merge(st.session_state.BD_Detalle_Tarimas, st.session_state.BD_Tarimas[['ID_Tarima', 'Estatus', 'Fecha_Creacion']], on='ID_Tarima', how='left')
        df_maestro_piezas['Estatus_Envio'] = df_maestro_piezas['Estatus'].apply(lambda x: "Remesado" if x == "Remesada" else "No Remesado")
        
        # Cuadrícula interactiva de 4 columnas en paralelo
        c_filt1, c_filt2, c_filt3, c_filt4 = st.columns(4)
        with c_filt1:
            opc_po = ["Todos"] + df_maestro_piezas['PO'].dropna().unique().tolist()
            f_po = st.selectbox("Orden de Compra (PO):", opc_po, key="query_po_filter_u")
            opc_sku = ["Todos"] + df_maestro_piezas['SKU'].dropna().unique().tolist()
            f_sku = st.selectbox("SKU / Producto:", opc_sku, key="query_sku_filter_u")
        with c_filt2:
            opc_proy = ["Todos"] + df_maestro_piezas['Project' if 'Project' in df_maestro_piezas.columns else 'Proyecto'].dropna().unique().tolist()
            f_proy = st.selectbox("Proyecto Interno:", opc_proy, key="query_proy_filter_u")
            opc_tar = ["Todos"] + df_maestro_piezas['ID_Tarima'].dropna().unique().tolist()
            f_tar = st.selectbox("ID Tarima:", opc_tar, key="query_tar_filter_u")
        with c_filt3:
            opc_parc = ["Todos"] + df_maestro_piezas['Parcialidad'].dropna().unique().tolist()
            f_parc = st.selectbox("Parcialidad:", opc_parc, key="query_parc_filter_u")
        with c_filt4:
            opc_desc = ["Todos"] + df_maestro_piezas['Descripcion'].dropna().unique().tolist()
            f_desc = st.selectbox("Descripción de Proyecto:", opc_desc, key="query_desc_filter_u")
            f_est = st.selectbox("Estatus de Envío:", ["Todos", "Remesado", "No Remesado"], key="query_est_filter_u")

        # Ejecución de la cascada de filtrado
        df_resultado = df_maestro_piezas.copy()
        if f_po != "Todos": df_resultado = df_resultado[df_resultado['PO'] == f_po]
        if f_sku != "Todos": df_resultado = df_resultado[df_resultado['SKU'] == f_sku]
        if f_proy != "Todos": df_resultado = df_resultado[df_resultado['Proyecto'] == f_proy]
        if f_tar != "Todos": df_resultado = df_resultado[df_resultado['ID_Tarima'] == f_tar]
        if f_parc != "Todos": df_resultado = df_resultado[df_resultado['Parcialidad'] == f_parc]
        if f_desc != "Todos": df_resultado = df_resultado[df_resultado['Descripcion'] == f_desc]
        if f_est != "Todos": df_resultado = df_resultado[df_resultado['Estatus_Envio'] == f_est]
        if not df_resultado.empty:
            df_rep = df_resultado[["ID_Tarima", "PO", "Proyecto", "Parcialidad", "Descripcion", "SKU", "Cantidad", "Estatus_Envio", "Fecha_Creacion"]].copy()
            df_rep.columns = ["ID Tarima", "Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", "Descripción de Proyecto Planta Rio", "SKU / Producto", "Cantidad (Pzs)", "Estatus de Envío", "Fecha de Ingreso"]
            
            total_piezas_consulta = int(df_rep['Cantidad (Pzs)'].sum())
            st.metric("🔢 Total Piezas en Selección:", f"{total_piezas_consulta:,} PZS")
            st.dataframe(df_rep, use_container_width=True, hide_index=True)
            
            # Construcción del reporte de auditoría multi-hoja con openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment
            from openpyxl.utils import get_column_letter
            
            buf_c = io.BytesIO()
            with pd.ExcelWriter(buf_c, engine='openpyxl') as writer_c:
                # Hoja 1: Portada institucional de criterios de control
                df_metadatos = pd.DataFrame([
                    {"Concepto": "DOCUMENTO", "Valor": "REPORTE CONSOLIDADO DE INVENTARIO POR PIEZA"},
                    {"Concepto": "EMPRESA", "Valor": "INDUSTRIA SIGRAMA S.A. DE C.V."},
                    {"Concepto": "FECHA DE GENERACIÓN", "Valor": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
                    {"Concepto": "FILTRO: ORDEN DE COMPRA (PO)", "Valor": str(f_po)},
                    {"Concepto": "FILTRO: PROYECTO INTERNO", "Valor": str(f_proy)},
                    {"Concepto": "FILTRO: PARCIALIDAD", "Valor": str(f_parc)},
                    {"Concepto": "FILTRO: DESCRIPCIÓN PROYECTO", "Valor": str(f_desc)},
                    {"Concepto": "FILTRO: SKU / PRODUCTO", "Valor": str(f_sku)},
                    {"Concepto": "FILTRO: ID TARIMA", "Valor": str(f_tar)},
                    {"Concepto": "FILTRO: ESTATUS DE ENVÍO", "Valor": str(f_est)},
                    {"Concepto": "TOTAL PIEZAS EN SELECCIÓN", "Valor": int(total_piezas_consulta)}
                ])
                df_metadatos.to_excel(writer_c, index=False, sheet_name='Resumen_Filtros')
                df_rep.to_excel(writer_c, index=False, sheet_name='Listado_Inventario')
                
                # Estilos de Portada (Gris de Control de Almacén)
                ws_m = writer_c.sheets['Resumen_Filtros']
                ws_m.row_dimensions.height = 24
                fill_m = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
                font_hm = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                for col_idx in range(1, 3):
                    cell = ws_m.cell(row=1, column=col_idx)
                    cell.fill, cell.font, cell.alignment = fill_m, font_hm, Alignment(horizontal="center", vertical="center")
                ws_m.column_dimensions['A'].width = 30
                ws_m.column_dimensions['B'].width = 50
                
                # Estilos de Lista de Datos (Rojo Sigrama Corporativo)
                ws_c = writer_c.sheets['Listado_Inventario']
                fill_h = PatternFill(start_color="D32F2F", end_color="D32F2F", fill_type="solid")
                font_h = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                ws_c.row_dimensions.height = 24
                for col_idx in range(1, len(df_rep.columns) + 1):
                    cell = ws_c.cell(row=1, column=col_idx)
                    cell.fill, cell.font, cell.alignment = fill_h, font_h, Alignment(horizontal="center", vertical="center")
                
                # Auto-ajuste seguro de anchos computado desde Pandas (Evita IndexError)
                for idx, col_name in enumerate(df_rep.columns):
                    max_len = max(df_rep[col_name].astype(str).map(len).max(), len(str(col_name)))
                    col_letter = get_column_letter(idx + 1)
                    ws_c.column_dimensions[col_letter].width = max(max_len + 4, 15)
                    
            buf_c.seek(0)
            st.download_button(label="📥 Generar Reporte de Inventario con Filtros y Fechas (.xlsx)", data=buf_c.getvalue(), file_name=f"Reporte_Inventario_Filtrado_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_download_consulta_piezas_excel_final_master")
        else:
            st.warning("⚠️ No existen registros que coincidan con los filtros seleccionados.")
    else:
        st.info("No hay datos de inventario registrados para realizar consultas por pieza.")
# =============================================================================
# 11. MÓDULO DE EMBAJALJE: CARGA MASIVA DE TARIMAS (DISEÑO CORPORATIVO REQUERIDO)
# =============================================================================
elif opcion_menu == "📦 Módulo Tarimas":
    st.title("📦 Carga de Tarimas")
    st.subheader("📋 Formato Requerido (Versión Proyectos y Parcialidades)")
    
    df_p = pd.DataFrame([{"Tarima": "Bulto_1", "Producto/SKU": "12-B-9016-01", "PO": "PO-10001", "Proyecto": "INT-001", "Parcialidad": "Entrega_P1", "Descripcion": "Proyecto Planta Rio", "Cantidad": 191}])
    
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    buf_p = io.BytesIO()
    with pd.ExcelWriter(buf_p, engine='openpyxl') as wr:
        df_p.to_excel(wr, index=False, sheet_name='Plantilla_Tarimas')
        workbook, worksheet = wr.book, wr.sheets['Plantilla_Tarimas']
        fill_rojo = PatternFill(start_color="D32F2F", end_color="D32F2F", fill_type="solid")
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        font_data = Font(name="Calibri", size=11, bold=False, color="000000")
        align_center, align_left = Alignment(horizontal="center", vertical="center"), Alignment(horizontal="left", vertical="center")
        borde_t = Side(border_style="thin", color="D3D3D3")
        borde_c = Border(left=borde_t, right=borde_t, top=borde_t, bottom=borde_t)
        
        worksheet.row_dimensions.height = 24
        for col_idx in range(1, 8):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill, cell.font, cell.alignment = fill_rojo, font_header, align_center
        worksheet.row_dimensions.height = 20
        for col_idx in range(1, 8):
            cell = worksheet.cell(row=2, column=col_idx)
            cell.font, cell.border = font_data, borde_c
            if col_idx in [1, 3, 4, 5, 7]:
                cell.alignment = align_center

            else: cell.alignment = align_left
                
        # Ajuste de anchos automático corregido leyendo la propiedad de la primera celda
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)  # <-- SOLUCIÓN ACÁ: col[0].column
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 15)

            
    buf_p.seek(0)
    st.download_button(label="📥 Descargar Formato de Plantilla Corporativa (.xlsx)", data=buf_p.getvalue(), file_name="plantilla_carga_tarimas_sigrama.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.write("---")
    
    if not st.session_state.BD_Tarimas.empty and "Es_Nueva" not in st.session_state.BD_Tarimas.columns:
        st.session_state.BD_Tarimas["Es_Nueva"] = False

    if not is_admin: st.error("🔒 Área Bloqueada: Se requiere contraseña de Administrador.")
    else:
        st.success("🔓 Acceso Autorizado.")
        arch = st.file_uploader("Suba el Excel con Formato de Proyectos", type=["xlsx"])
        col_t1, col_t2 = st.columns(2)
        with col_t1: tipo_t = st.selectbox("Tipo:", ["Cuadrada", "Rectangular"])
        with col_t2: oper = st.text_input("Líder:", "Jesus Morales")
        if arch and st.button("Procesar e Integrar Plantilla Avanzada"):
            try:
                df_ex = pd.read_excel(arch)
                columnas_requeridas = ["Tarima", "Producto/SKU", "PO", "Proyecto", "Parcialidad", "Descripcion", "Cantidad"]
                if not all(col in df_ex.columns for col in columnas_requeridas): st.error("❌ Error: Columnas incompatibles.")
                else:
                    if not st.session_state.BD_Tarimas.empty: st.session_state.BD_Tarimas["Es_Nueva"] = False
                    for t_orig in df_ex['Tarima'].unique():
                        nuevo_id_tpm = f"TPM-{(len(st.session_state.BD_Tarimas) + 1):04d}"
                        n_t = {"ID_Tarima": nuevo_id_tpm, "Tarima_Origen_Excel": t_orig, "Fecha_Creacion": datetime.datetime.now().strftime("%d/%m/%Y"), "Ubicacion_Actual": "Metales", "Creado_Por": oper, "Tipo_Tarima": tipo_t, "Estatus": "Disponible", "Es_Nueva": True}
                        st.session_state.BD_Tarimas = pd.concat([st.session_state.BD_Tarimas, pd.DataFrame([n_t])], ignore_index=True)
                        items = df_ex[df_ex['Tarima'] == t_orig]
                        for _, item in items.iterrows():
                            st.session_state.BD_Detalle_Tarimas = pd.concat([st.session_state.BD_Detalle_Tarimas, pd.DataFrame([{"ID_Detalle": len(st.session_state.BD_Detalle_Tarimas) + 1, "ID_Tarima": nuevo_id_tpm, "SKU": item['Producto/SKU'], "PO": item['PO'], "Proyecto": item['Proyecto'], "Parcialidad": item['Parcialidad'], "Descripcion": item['Descripcion'], "Cantidad": item['Cantidad']}])], ignore_index=True)
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                    st.success("¡Inventario respaldado con éxito!"); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
            
    if not st.session_state.BD_Tarimas.empty:
        st.write("---")
        st.subheader("🖨️ Panel de Impresión Masiva de Tarimas")
        df_estilado = st.session_state.BD_Tarimas.style.apply(lambda r: ['background-color: #FFF59D' if r['Es_Nueva'] else '' for _ in r], axis=1)
        seleccion_tabla = st.dataframe(df_estilado, use_container_width=True, column_order=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus"], on_select="rerun", selection_mode="multi-row")
        filas_seleccionadas = seleccion_tabla.get("selection", {}).get("rows", [])
        if filas_seleccionadas:
            elegidas = st.session_state.BD_Tarimas.iloc[filas_seleccionadas]['ID_Tarima'].tolist()
            if len(elegidas) == 1:
                st.download_button(label=f"📥 Descargar PDF Tarima #{elegidas}", data=generar_pdf_tarima(elegidas), file_name=f"Tarima_{elegidas}.pdf", mime="application/pdf")
            else:
                if st.button("📦 Unificar Lote de Impresión"):
                    buf_l = io.BytesIO()
                    doc_l = SimpleDocTemplate(buf_l, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
                    story_l, styles = [], getSampleStyleSheet()
                    for t_imp in elegidas:
                        det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == t_imp]
                        t_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == t_imp].iloc
                        story_l.append(Spacer(1, 1.8 * inch)); story_l.append(Paragraph(f"TARIMA<br/><br/><b>#{t_imp}</b>", ParagraphStyle('G_L', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1))); story_l.append(PageBreak())
                        story_l.append(Paragraph(f"<b>Detalle Interno - Tarima #{t_imp}</b>", styles['Heading2']))
                        story_l.append(Paragraph(f"<b>Operador:</b> {t_info['Creado_Por']} | <b>Fecha:</b> {t_info['Fecha_Creacion']}", styles['Normal']))
                        story_l.append(Spacer(1, 0.3 * inch))
                        for _, item in det.iterrows():
                            art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
                            story_l.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {art.iloc['Nombre'] if not art.empty else 'Material'}", styles['Normal']))
                            story_l.append(Spacer(1, 0.4 * inch)); story_l.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", ParagraphStyle('NG_L', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)))
                        story_l.append(PageBreak())
                    if story_l: story_l.pop()
                    doc_l.build(story_l, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
                    st.download_button(label="📥 Descargar Lote Completo (PDF)", data=buf_l.getvalue(), file_name="Lote_Tarimas.pdf", mime="application/pdf")
        else: st.warning("Seleccione registros en la tabla usando las casillas.")
# =============================================================================
# 12. MÓDULO DE DESPACHOS LOGÍSTICOS: EMISIÓN DE REMISIONES DE SALIDA
# =============================================================================
elif opcion_menu == "🚚 Módulo Remisiones":
    st.title("🚚 Generación de Remisiones de Salida")
    t_disp = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible']['ID_Tarima'].tolist()
    if not t_disp: st.warning("⚠️ No existen tarimas disponibles en patio.")
    else:
        t_sel = st.multiselect("Seleccione Tarimas para este Envío:", options=t_disp)
        col_e, col_r = st.columns(2)
        with col_e:
            # Lista desplegable acoplada de forma dinámica al catálogo relacional
            if "BD_Lideres" in st.session_state and not st.session_state.BD_Lideres.empty:
                lista_nombres_lideres = st.session_state.BD_Lideres['Nombre_Lider'].unique().tolist()
                nom_e = st.selectbox("Líder / Emisor Autorizado:", options=lista_nombres_lideres, key="rem_lider_sel_unique")
            else:
                nom_e = st.selectbox("Líder / Emisor Autorizado:", options=["Jesus Morales", "Supervisor General"], key="rem_lider_backup_unique")
            dir_e = st.text_input("Almacén de Origen:", "Metales")
        with col_r:
            nom_r = st.text_input("Receptor / Cliente:", "Galvatec Industrias")
            dir_r = st.text_input("Dirección Destino:", "Prol. Valle Guadiana 919, Parque Industrial II, 35078 Gómez Palacio, Dgo.")
        if not is_admin: st.error("🔒 Operación Bloqueada: Se requiere contraseña de Administrador.")
        else:
            if st.button("🚀 Confirmar Salida y Generar Nueva Remisión"):
                if not t_sel or not nom_r: st.error("Complete los campos obligatorios.")
                else:
                    fol = f"E00{27 + len(st.session_state.BD_Datos_Generales_Remision)}"
                    reg = {"ID_Remision": len(st.session_state.BD_Datos_Generales_Remision)+1, "Folio_Remision": fol, "Fecha_Hora_Salida": datetime.datetime.now().strftime("%d/%m/%Y"), "Nombre_Emisor": nom_e, "Direccion_Emisor": dir_e, "Nombre_Receptor": nom_r, "Direccion_Receptor": dir_r, "Tarimas_Asociadas": t_sel}
                    st.session_state.BD_Datos_Generales_Remision = pd.concat([st.session_state.BD_Datos_Generales_Remision, pd.DataFrame([reg])], ignore_index=True)
                    st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(t_sel), 'Estatus'] = 'Remesada'
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success(f"✅ ¡Remisión {fol} Generada y Guardada de Forma Permanente!"); st.rerun()
                    
    if not st.session_state.BD_Datos_Generales_Remision.empty:
        st.write("---")
        st.subheader("🖨️ Descarga Documental de Remisiones")
        r_sel = st.selectbox("Seleccione Folio para Descarga:", st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].unique(), key="rem_download_folio_sel")

        # SOLUCIÓN DE RAÍZ: Extraer la fila como un diccionario nativo de Python para eliminar indexadores de Pandas
        lista_remisiones = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'] == r_sel].to_dict('records')
        
        if lista_remisiones:
            row_dict = lista_remisiones[0]
            df_det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(row_dict['Tarimas_Asociadas'])]
            
            c1, c2 = st.columns(2)
            with c1: st.download_button("📥 Descargar Remisión (PDF)", data=generar_pdf_remision_general(row_dict, df_det), file_name=f"Remision_{r_sel}.pdf", key="btn_dl_rem_pdf", mime="application/pdf")
            with c2: st.download_button("📥 Descargar Anexo Tarimas (PDF)", data=generar_pdf_anexo_tarimas(row_dict['Tarimas_Asociadas'], df_det), file_name=f"Anexo_{r_sel}.pdf", key="btn_dl_anexo_pdf", mime="application/pdf")




# =============================================================================
# SECCIÓN 17A: PANEL DE MANTENIMIENTO AVANZADO Y EDICIÓN EN CALIENTE (SUPERUSER)
# =============================================================================
elif opcion_menu == "⚙️ Mantenimiento y Catálogos":
    st.title("⚙️ Panel de Mantenimiento Avanzado del Sistema")
    st.warning("⚠️ Acción Crítica: Las modificaciones realizadas impactan directamente en los archivos de GitHub.")
    
    # Inicialización del catálogo básico de personal operativo en memoria
    if "BD_Lideres" not in st.session_state:
        st.session_state.BD_Lideres = pd.DataFrame([
            {"ID_Lider": "LID-01", "Nombre_Lider": "Jesus Morales", "Area": "Metales", "Estatus": "Activo"}
        ])

    tab1, tab2, tab3 = st.tabs(["📝 Ajustar Cantidades", "👤 Catálogo de Líderes", "🚨 Purga de Datos"])

    # --- SUB-MÓDULO 1: MODIFICACIÓN DIRECTA DE CANTIDADES DE MATERIALES ---
    with tab1:
        st.subheader("✏️ Edición Rápida de Inventario (Detalle Tarimas)")
        if not st.session_state.BD_Detalle_Tarimas.empty:
            st.markdown("Haga doble clic sobre la celda de la columna **Cantidad** para corregir errores de dedo:")
            
            # Editor interactivo tipo Excel bloqueando llaves relacionales para evitar inconsistencias
            df_editable = st.data_editor(
                st.session_state.BD_Detalle_Tarimas, 
                use_container_width=True, 
                disabled=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Proyecto", "Parcialidad", "Descripcion"], 
                hide_index=True, 
                key="editor_mantenimiento_cantidades_final"
            )
            
            if st.button("💾 Guardar Cambios de Cantidades en GitHub"):
                st.session_state.BD_Detalle_Tarimas = df_editable
                subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                st.success("✅ ¡Cantidades corregidas y sincronizadas con éxito en el repositorio!"); st.rerun()
        else: 
            st.info("No existen registros en el desglose de inventario para modificar.")
# =============================================================================
# SECCIÓN 17B: MANTENIMIENTO - CATALOGO DE LÍDERES MEDIANTE ARCHIVOS MASIVOS
# =============================================================================
    with tab2:
        st.subheader("👤 Administración del Personal de Líderes")
        
        # Generación de archivo muestra descargable para los operadores de IT
        df_l_plantilla = pd.DataFrame([
            {"Nombre_Lider": "Nombre Ejemplo 1", "Area": "Metales"},
            {"Nombre_Lider": "Nombre Ejemplo 2", "Area": "Embarques"}
        ])
        
        from openpyxl.styles import Font, PatternFill
        buf_l = io.BytesIO()
        with pd.ExcelWriter(buf_l, engine='openpyxl') as wr_l:
            df_l_plantilla.to_excel(wr_l, index=False, sheet_name='Plantilla_Lideres')
            ws_l = wr_l.sheets['Plantilla_Lideres']
            
            fill_header_lider = PatternFill(start_color="757575", end_color="757575", fill_type="solid")
            font_header_lider = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            for col_idx in range(1, 3):
                cell = ws_l.cell(row=1, column=col_idx)
                cell.font, cell.fill = font_header_lider, fill_header_lider
                
        st.write("Descargue el formato base para rellenar la lista de personal autorizado:")
        st.download_button(label="📥 Descargar Plantilla de Personal (.xlsx)", data=buf_l.getvalue(), file_name="plantilla_lideres_sigrama.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_download_plantilla_lideres_u")
        st.write("---")
        
        # Procesador de integración masiva de empleados
        arch_lideres = st.file_uploader("Suba la Plantilla de Personal Rellenada:", type=["xlsx"], key="uploader_lideres_masivo_u")
        if arch_lideres and st.button("🚀 Procesar e Integrar Personal Masivo"):
            try:
                df_l_excel = pd.read_excel(arch_lideres)
                if not all(col in df_l_excel.columns for col in ["Nombre_Lider", "Area"]):
                    st.error("❌ Error: Columnas incompatibles. Use: Nombre_Lider, Area")
                else:
                    for _, row_l in df_l_excel.iterrows():
                        if pd.notna(row_l['Nombre_Lider']) and str(row_l['Nombre_Lider']).strip() != "":
                            n_id_l = f"LID-{(len(st.session_state.BD_Lideres) + 1):02d}"
                            n_l_row = {"ID_Lider": n_id_l, "Nombre_Lider": str(row_l['Nombre_Lider']).strip(), "Area": str(row_l['Area']).strip() if pd.notna(row_l['Area']) else "Metales", "Estatus": "Activo"}
                            st.session_state.BD_Lideres = pd.concat([st.session_state.BD_Lideres, pd.DataFrame([n_l_row])], ignore_index=True)
                    subir_excel_a_github("BD_Lideres.xlsx", st.session_state.BD_Lideres)
                    st.success("✅ ¡Lista de personal integrada y respaldada con éxito!"); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
# =============================================================================
# SECCIÓN 17C: MANTENIMIENTO - REGISTRO TRADICIONAL Y RESPALDO COMPACTO
# =============================================================================
        st.write("---")
        with st.expander("➕ Dar de Alta Nuevo Líder Individual"):
            c_l1, c_l2 = st.columns(2)
            with c_l1: nuevo_nom = st.text_input("Nombre Completo del Líder:", key="txt_input_nuevo_lider_name_f")
            with c_l2: nueva_area = st.text_input("Área de Adscripción:", "Metales", key="txt_input_nuevo_lider_area_f")
            if st.button("➕ Registrar Líder Individual"):
                if nuevo_nom:
                    n_row = {"ID_Lider": f"LID-{(len(st.session_state.BD_Lideres) + 1):02d}", "Nombre_Lider": nuevo_nom, "Area": nueva_area, "Estatus": "Activo"}
                    st.session_state.BD_Lideres = pd.concat([st.session_state.BD_Lideres, pd.DataFrame([n_row])], ignore_index=True)
                    subir_excel_a_github("BD_Lideres.xlsx", st.session_state.BD_Lideres)
                    st.success("Líder registrado exitosamente."); st.rerun()
                    
        st.write("📋 Catálogo Maestro de Líderes Autorizados (Editable):")
        df_l_edit = st.data_editor(st.session_state.BD_Lideres, use_container_width=True, hide_index=True, key="editor_catalogo_lideres_master_final")
        if st.button("💾 Sincronizar Cambios Manuales de Líderes"):
            st.session_state.BD_Lideres = df_l_edit
            subir_excel_a_github("BD_Lideres.xlsx", st.session_state.BD_Lideres)
            st.success("Catálogo de personal operativo actualizado en GitHub.")
# =============================================================================
# SECCIÓN 17D: MANTENIMIENTO - PURGA SELECCIONADA Y CIERRE GENERAL DEL SISTEMA
# =============================================================================
    with tab3:
        st.subheader("🚨 Reset de Fábrica y Purga de Datos Controlada")
        metodo_purga = st.radio("Método de Purga:", ["❌ Purga Total Automática (Reset Completo)", "🔍 Seleccionar Registros Específicos para Eliminar"], horizontal=True, key="radio_metodo_purga_master_final")
        st.write("---")
        
        if metodo_purga == "❌ Purga Total Automática (Reset Completo)":
            st.error("⚠️ Peligro: Esta acción vaciará por completo todos los históricos en el repositorio.")
            if st.checkbox("Confirmo que deseo aplicar un Reset de Fábrica Total.", key="chk_total_purga_final_final"):
                if st.button("🗑️ EJECUTAR PURGA MAESTRA TOTAL"):
                    st.session_state.BD_Tarimas = pd.DataFrame(columns=st.session_state.BD_Tarimas.columns)
                    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=st.session_state.BD_Detalle_Tarimas.columns)
                    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=st.session_state.BD_Datos_Generales_Remision.columns)
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success("💥 Sistema purgado por completo a ceros de forma masiva."); st.rerun()
        else:
            st.markdown("### 📦 1. Eliminar Tarimas del Inventario")
            if not st.session_state.BD_Tarimas.empty:
                df_tar_vista = st.session_state.BD_Tarimas.copy().drop(columns=["Es_Nueva"], errors="ignore")
                sel_tarimas = st.dataframe(df_tar_vista, use_container_width=True, on_select="rerun", selection_mode="multi-row", key="tabla_purga_tarimas_final_f")
                filas_tar = sel_tarimas.get("selection", {}).get("rows", [])
                if filas_tar:
                    ids_tar_eliminar = st.session_state.BD_Tarimas.iloc[filas_tar]['ID_Tarima'].tolist()
                    if st.button("🗑️ Eliminar Tarimas Seleccionadas"):
                        st.session_state.BD_Tarimas = st.session_state.BD_Tarimas[~st.session_state.BD_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                        st.session_state.BD_Detalle_Tarimas = st.session_state.BD_Detalle_Tarimas[~st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                        subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                        subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                        st.success("✅ Tarimas removidas."); st.rerun()
            else: st.write("No hay tarimas registradas.")
            
            st.write("---")
            st.markdown("### 🚚 2. Eliminar Remisiones de Salida")
            if not st.session_state.BD_Datos_Generales_Remision.empty:
                sel_remisiones = st.dataframe(st.session_state.BD_Datos_Generales_Remision, use_container_width=True, on_select="rerun", selection_mode="multi-row", key="tabla_purga_remisiones_final_f")
                filas_rem = sel_remisiones.get("selection", {}).get("rows", [])
                if filas_rem:
                    ids_rem_eliminar = st.session_state.BD_Datos_Generales_Remision.iloc[filas_rem]['Folio_Remision'].tolist()
                    tarimas_afectadas = []
                    for idx in filas_rem: tarimas_afectadas.extend(st.session_state.BD_Datos_Generales_Remision.iloc[idx]['Tarimas_Asociadas'])
                    if st.button("🗑️ Eliminar Remisiones Seleccionadas"):
                        st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(tarimas_afectadas), 'Estatus'] = 'Disponible'
                        st.session_state.BD_Datos_Generales_Remision = st.session_state.BD_Datos_Generales_Remision[~st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].isin(ids_rem_eliminar)]
                        subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                        subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                        st.success("✅ Remisiones eliminadas y bultos reactivados."); st.rerun()
            else: st.write("No hay remisiones registradas.")
