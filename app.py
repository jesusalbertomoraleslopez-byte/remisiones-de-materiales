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

# 1. CONFIGURACIÓN DE LA PÁGINA WEB RESPONSIVA
st.set_page_config(page_title="Remisiones de Materiales", layout="wide", page_icon="📦")

# BANNER SUPERIOR CORPORATIVO ADAPTABLE
try:
    banner_img = Image.open("REMISIONES APP.png")
    st.image(banner_img, use_container_width=True)
except FileNotFoundError:
    st.warning("⚠️ Cargando interfaz gráfica del banner...")
st.write("")

# 2. MODELO RELACIONAL CON PROYECTOS Y PARCIALIDADES EN SESSION STATE
if "BD_Articulos" not in st.session_state:
    st.session_state.BD_Articulos = pd.DataFrame([
        {"SKU": "12-B-9016-01", "Nombre": "Lámina Galvanizada Sigrama", "Calibre_Espesor": "Calibre 22", "Dimensiones_Pieza": "3x10 ft", "Acabado_Superficial": "Zintro"},
        {"SKU": "SKU-002", "Nombre": "Placa de Acero Comercial", "Calibre_Espesor": "1/4 pulgada", "Dimensiones_Pieza": "4x8 ft", "Acabado_Superficial": "Negro"}
    ])

if "BD_Tarimas" not in st.session_state:
    st.session_state.BD_Tarimas = pd.DataFrame(columns=[
        "ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus", "Es_Nueva"
    ])

if "BD_Detalle_Tarimas" not in st.session_state:
    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=[
        "ID_Detalle", "ID_Tarima", "SKU", "PO", "Proyecto", "Parcialidad", "Descripcion", "Cantidad"
    ])

if "BD_Datos_Generales_Remision" not in st.session_state:
    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=[
        "ID_Remision", "Folio_Remision", "Fecha_Hora_Salida", "Nombre_Emisor", "Direccion_Emisor", "Nombre_Receptor", "Direccion_Receptor", "Tarimas_Asociadas"
    ])
# 3. CONTROL DE SEGURIDAD MÚLTIPLE (ADMINISTRADOR Y SUPERUSUARIO)
st.sidebar.title("🔐 Control de Acceso")
admin_pass_input = st.sidebar.text_input("Contraseña Administrador:", type="password")

def es_admin():
    try:
        return admin_pass_input == st.secrets["admin_password"]
    except KeyError:
        st.sidebar.error("Error: 'admin_password' no configurado en los Secrets de la plataforma.")
        return False

is_admin = es_admin()
if is_admin:
    st.sidebar.success("Modo Administrador Activo")
else:
    st.sidebar.warning("Modo Consulta Activo")

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
# 4. CONECTOR DE PERSISTENCIA AUTOMÁTICA EN GITHUB
def subir_excel_a_github(file_name, dataframe_to_save):
    try:
        GITHUB_TOKEN = st.secrets["github_token"]
        REPO_OWNER = "jesusalbertomoraleslopez-byte"
        REPO_NAME = "remisiones-de-materiales"
        BRANCH = "main"
        
        buffer_git = io.BytesIO()
        with pd.ExcelWriter(buffer_git, engine='openpyxl') as writer:
            dataframe_to_save.to_excel(writer, index=False, sheet_name='Datos_Sistema')
        content_bytes = buffer_git.getvalue()
        base64_content = base64.b64encode(content_bytes).decode("utf-8")
        
        # URL CORREGIDA DE LA API OFICIAL DE GITHUB
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{file_name}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
        
        res_get = requests.get(url, headers=headers)
        sha = None
        if res_get.status_code == 200:
            sha = res_get.json().get("sha")
            
        payload = {"message": f"Sincronización App Streamlit: {file_name}", "content": base64_content, "branch": BRANCH}
        if sha:
            payload["sha"] = sha
            
        res_put = requests.put(url, json=payload, headers=headers)
        if res_put.status_code in [200, 201]:
            st.sidebar.success(f"💾 Respaldo en Git: {file_name}")
            return True
        else:
            return False
    except Exception as e:
        return False
# 5. MAQUETACIÓN DOCUMENTAL INSTITUCIONAL SIGRAMA (FO-MET-10)
def draw_sigrama_decorations(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#D32F2F"))
    canvas.rect(36, 745, 540, 4, fill=1, stroke=0)
    
    canvas.setFont("Helvetica-Bold", 11)
    canvas.setFillColor(colors.HexColor("#D32F2F"))
    canvas.drawString(36, 765, "FO-MET-10")
    
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.black)
    canvas.drawString(36, 753, "Revisión 01")
    canvas.drawString(36, 741, "04 de octubre 2018")
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawCentredString(285, 755, "EMBARQUE-RECEPCIÓN DE MERCANCÍA")
    canvas.setStrokeColor(colors.HexColor("#D32F2F"))
    canvas.line(36, 45, 36, 25)
    canvas.setFont("Helvetica-Bold", 7)
    canvas.drawString(42, 37, "FO-SGC-02")
    canvas.setFont("Helvetica", 6)
    canvas.setFillColor(colors.HexColor("#424242"))
    texto_legal = "PROHIBIDA LA REPRODUCCIÓN TOTAL O PARCIAL, POR CUALQUIER MEDIO O PROCEDIMIENTO, SIN AUTORIZACIÓN DE INDUSTRIA SIGRAMA S.A. DE C.V."
    canvas.drawString(95, 37, texto_legal)
    canvas.restoreState()
def generar_pdf_tarima(id_tarima):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    tarima_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == id_tarima].iloc[0]
    detalles = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == id_tarima]
    
    # Hoja 1: Carátula gigante
    style_g = ParagraphStyle('G', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1)
    story.append(Spacer(1, 1.8 * inch))
    story.append(Paragraph(f"TARIMA<br/><br/><b>#{id_tarima}</b>", style_g))
    story.append(PageBreak())
    
    # Hoja 2: Detalles técnicos de bulto
    style_n = styles['Normal']
    style_ng = ParagraphStyle('NG', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)
    story.append(Paragraph(f"<b>Detalle Interno - Tarima #{id_tarima}</b>", styles['Heading2']))
    story.append(Paragraph(f"<b>Líder Operativo:</b> {tarima_info['Creado_Por']} | <b>Fecha:</b> {tarima_info['Fecha_Creacion']}", style_n))
    story.append(Spacer(1, 0.3 * inch))
    
    for _, item in detalles.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
        story.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", style_n))
        if 'Proyecto' in item:
            story.append(Paragraph(f"<b>Proyecto:</b> {item['Proyecto']} | <b>Parcialidad:</b> {item['Parcialidad']}", style_n))
        story.append(Spacer(1, 0.4 * inch))
        story.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", style_ng))
        
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer.getvalue()
def generar_pdf_remision_general(datos_remision, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    style_blanco_bold = ParagraphStyle('BB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_normal_bold = ParagraphStyle('NB', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=9)
    style_normal_text = ParagraphStyle('NT', parent=styles['Normal'], fontSize=9)
    
    story.append(Spacer(1, 0.1 * inch))
    t_header = Table([[Paragraph("LISTADO DE EMBARQUE", style_blanco_bold)]], colWidths=[540])
    t_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header)
    
    datos_panel = [
        [Paragraph("FOLIO:", style_normal_bold), Paragraph(str(datos_remision['Folio_Remision']), style_normal_text), Paragraph("FECHA:", style_normal_bold), Paragraph(str(datos_remision['Fecha_Hora_Salida']), style_normal_text)],
        [Paragraph("LÍDER:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Emisor']), style_normal_text), Paragraph("ALMACÉN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Emisor']), style_normal_text)],
        [Paragraph("DESTINO:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Receptor']), style_normal_text), "", ""],
        [Paragraph("DIRECCIÓN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Receptor']), style_normal_text), "", ""]
    ]
    t_panel = Table(datos_panel, colWidths=[100, 170, 100, 170])
    t_panel.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")), ('SPAN', (1,2), (3,2)), ('SPAN', (1,3), (3,3))]))
    story.append(t_panel)
    story.append(Spacer(1, 0.2 * inch))
    
    tabla_materiales = [[Paragraph("CANTIDAD", style_blanco_bold), Paragraph("UNIDAD", style_blanco_bold), Paragraph("CLAVE / MODELO", style_blanco_bold), Paragraph("DESCRIPCIÓN", style_blanco_bold), Paragraph("OBSERVACIONES", style_blanco_bold)]]
    for _, row in df_detalles_remision.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == row['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Material"
        tabla_materiales.append([Paragraph(str(int(row['Cantidad'])), style_normal_text), Paragraph("Piezas", style_normal_text), Paragraph(str(row['SKU']), style_normal_text), Paragraph(nom_art, style_normal_text), Paragraph(f"Tarima: {row['ID_Tarima']}", style_normal_text)])
    
    for _ in range(max(1, 6 - len(df_detalles_remision))):
        tabla_materiales.append([Paragraph("", style_normal_text)] * 5)
        
    t_mat = Table(tabla_materiales, colWidths=[70, 60, 110, 180, 120])
    t_mat.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575"))]))
    story.append(t_mat)
    
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer.getvalue()

def generar_pdf_anexo_tarimas(lista_tarimas_id, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    style_b = ParagraphStyle('ABB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_t = ParagraphStyle('ANT', parent=styles['Normal'], fontSize=9)
    
    # LEADING INTEGRADO PARA EVITAR EMPALMES VISUALES
    style_caratula = ParagraphStyle('C_Fix', parent=styles['Heading1'], fontSize=42, leading=46, alignment=1)
    
    for t_id in lista_tarimas_id:
        story.append(Spacer(1, 1.8 * inch))
        story.append(Paragraph(f"ANEXO: TARIMA<br/><b>{t_id}</b>", style_caratula))
        story.append(Spacer(1, 0.4 * inch))
        t_bar = Table([["|||||||||||||||||||||||||||||||"], [f"*{t_id}*"]], colWidths=[468])
        t_bar.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('TEXTCOLOR', (0,0), (-1,1), colors.darkgray)]))
        story.append(t_bar)
        story.append(PageBreak())
        
        story.append(Paragraph(f"<b>DETALLE ESPECÍFICO - TARIMA {t_id}</b>", styles['Heading2']))
        sub_det = df_detalles_remision[df_detalles_remision['ID_Tarima'] == t_id]
        tabla_anexo = [[Paragraph("PO ASOCIADA", style_b), Paragraph("SKU / PRODUCTO", style_b), Paragraph("CANTIDAD", style_b)]]
        for _, s_row in sub_det.iterrows():
            tabla_anexo.append([Paragraph(str(s_row['PO']), style_t), Paragraph(str(s_row['SKU']), style_t), Paragraph(str(int(s_row['Cantidad'])), style_t)])
        t_det = Table(tabla_anexo, colWidths=[156, 156, 156])
        t_det.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F5F5F5")), ('GRID', (0,0), (-1,-1), 1, colors.white)]))
        story.append(t_det)
        story.append(PageBreak())
        
    if story: story.pop()
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer.getvalue()
# 6. RENDERIZADO DE CAPA VISUAL
if opcion_menu == "📊 Dashboard e Históricos":
    st.title("📊 Dashboard Planta Metales Inventario Producto")
    col_f1, col_f2 = st.columns(2)
    with col_f1: f_inicio = st.date_input("Fecha Inicial", datetime.date.today() - datetime.timedelta(days=7))
    with col_f2: f_fin = st.date_input("Fecha Final", datetime.date.today())
    
    t_tar = len(st.session_state.BD_Tarimas)
    disp = len(st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible'])
    rem = len(st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Remesada'])
    
    if not st.session_state.BD_Detalle_Tarimas.empty:
        total_piezas = int(st.session_state.BD_Detalle_Tarimas['Cantidad'].sum())
        id_disp = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible']['ID_Tarima']
        piezas_disp = int(st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(id_disp)]['Cantidad'].sum())
        piezas_rem = total_piezas - piezas_disp
    else:
        total_piezas = piezas_disp = piezas_rem = 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📦 Total Tarimas", f"{t_tar} Reg.")
    m2.metric("🟢 Tarimas Disponibles", disp)
    m3.metric("🚚 Tarimas Remesadas", rem)
    m4.metric("🔢 Total Piezas", f"{total_piezas:,} PZS")
    
    st.write("---")
    st.subheader("🔍 Filtros Operativos del Proyecto")
    col_x1, col_x2, col_x3 = st.columns(3)
    df_c_filtro = st.session_state.BD_Detalle_Tarimas.copy()
    
    with col_x1:
        opciones_proy = ["Todos"] + df_c_filtro['Proyecto'].dropna().unique().tolist() if not df_c_filtro.empty and 'Proyecto' in df_c_filtro.columns else ["Todos"]
        proy_sel = st.selectbox("Filtrar por Proyecto Interno:", opciones_proy)
    with col_x2:
        if proy_sel != "Todos": df_c_filtro = df_c_filtro[df_c_filtro['Proyecto'] == proy_sel]
        opciones_po = ["Todas"] + df_c_filtro['PO'].dropna().unique().tolist() if not df_c_filtro.empty else ["Todas"]
        po_sel = st.selectbox("Filtrar por Orden de Compra (PO):", opciones_po)
    with col_x3:
        if po_sel != "Todas": df_c_filtro = df_c_filtro[df_c_filtro['PO'] == po_sel]
        opciones_parc = ["Todas"] + df_c_filtro['Parcialidad'].dropna().unique().tolist() if not df_c_filtro.empty and 'Parcialidad' in df_c_filtro.columns else ["Todas"]
        parc_sel = st.selectbox("Filtrar por Parcialidad:", opciones_parc)
        
    st.write("---")
    st.subheader("📋 Resumen de Cumplimiento por Pedido (PO) y Parcialidades")
    if not st.session_state.BD_Detalle_Tarimas.empty:
        df_completo = pd.merge(st.session_state.BD_Detalle_Tarimas, st.session_state.BD_Tarimas[['ID_Tarima', 'Estatus']], on='ID_Tarima', how='left')
        if proy_sel != "Todos": df_completo = df_completo[df_completo['Proyecto'] == proy_sel]
        if po_sel != "Todas": df_completo = df_completo[df_completo['PO'] == po_sel]
        if parc_sel != "Todas": df_completo = df_completo[df_completo['Parcialidad'] == parc_sel]
            
        if not df_completo.empty and all(col in df_completo.columns for col in ['Proyecto', 'Parcialidad', 'Descripcion']):
            resumen_avanzado = df_completo.groupby(['PO', 'Proyecto', 'Parcialidad', 'Descripcion']).agg(
                Cant_Tarimas=('ID_Tarima', 'nunique'), Total_Piezas=('Cantidad', 'sum'),
                Piezas_Disponibles=('Cantidad', lambda x: x[df_completo.loc[x.index, 'Estatus'] == 'Disponible'].sum()),
                Piezas_Remesadas=('Cantidad', lambda x: x[df_completo.loc[x.index, 'Estatus'] == 'Remesada'].sum())
            ).reset_index()
            resumen_avanzado['% Avance Salida'] = (resumen_avanzado['Piezas_Remesadas'] / resumen_avanzado['Total_Piezas'] * 100).round(1)
            resumen_avanzado['% Avance Salida'] = resumen_avanzado['% Avance Salida'].apply(lambda x: f"{x}%")
            resumen_avanzado.columns = ["Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", "Descripción del Lote", "Cant. Tarimas", "Total Piezas", "Piezas Disponibles", "Piezas Remesadas", "% Avance Salida"]
            st.dataframe(resumen_avanzado, use_container_width=True, hide_index=True)
    else:
        st.info("No hay datos de proyectos o parcialidades cargados.")
def generar_pdf_remision_general(datos_remision, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    style_blanco_bold = ParagraphStyle('BB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_normal_bold = ParagraphStyle('NB', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=9)
    style_normal_text = ParagraphStyle('NT', parent=styles['Normal'], fontSize=9)
    
    story.append(Spacer(1, 0.1 * inch))
    t_header_embarque = Table([[Paragraph("LISTADO DE EMBARQUE", style_blanco_bold)]], colWidths=[540])
    t_header_embarque.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header_embarque)
    datos_panel = [
        [Paragraph("FOLIO:", style_normal_bold), Paragraph(str(datos_remision['Folio_Remision']), style_normal_text), Paragraph("FECHA:", style_normal_bold), Paragraph(str(datos_remision['Fecha_Hora_Salida']), style_normal_text)],
        [Paragraph("LÍDER:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Emisor']), style_normal_text), Paragraph("NO. ALMACÉN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Emisor']), style_normal_text)],
        [Paragraph("DESTINO:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Receptor']), style_normal_text), "", ""],
        [Paragraph("DIRECCIÓN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Receptor']), style_normal_text), "", ""]
    ]
    t_panel = Table(datos_panel, colWidths=[110, 160, 100, 170])
    t_panel.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")), ('SPAN', (1,2), (3,2)), ('SPAN', (1,3), (3,3))]))
    story.append(t_panel)
    story.append(Spacer(1, 0.2 * inch))
    
    tabla_materiales = [[Paragraph("CANTIDAD", style_blanco_bold), Paragraph("UNIDAD", style_blanco_bold), Paragraph("CLAVE / MODELO", style_blanco_bold), Paragraph("DESCRIPCIÓN", style_blanco_bold), Paragraph("OBSERVACIONES", style_blanco_bold)]]
    for _, row in df_detalles_remision.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == row['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Material"
        tabla_materiales.append([Paragraph(str(int(row['Cantidad'])), style_normal_text), Paragraph("Piezas", style_normal_text), Paragraph(str(row['SKU']), style_normal_text), Paragraph(nom_art, style_normal_text), Paragraph(f"Tarima: {row['ID_Tarima']}", style_normal_text)])
    for _ in range(max(1, 8 - len(df_detalles_remision))):
        tabla_materiales.append([Paragraph("", style_normal_text)] * 5)
        
    t_mat = Table(tabla_materiales, colWidths=[70, 60, 110, 180, 120])
    t_mat.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575"))]))
    story.append(t_mat)
    story.append(Spacer(1, 0.3 * inch))
    t_firmas = Table([[Paragraph("ENTREGA", style_normal_bold), Paragraph("RECIBE", style_normal_bold)]], colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('LINEBELOW', (0,0), (0,0), 1, colors.black), ('LINEBELOW', (1,0), (1,0), 1, colors.black), ('TOPPADDING', (0,0), (-1,-1), 25)]))
    story.append(t_firmas)
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer
def generar_pdf_anexo_tarimas(lista_tarimas_id, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    style_b = ParagraphStyle('ABB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_t = ParagraphStyle('ANT', parent=styles['Normal'], fontSize=9)
    style_caratula = ParagraphStyle('C_Fix', parent=styles['Heading1'], fontSize=42, leading=46, alignment=1)
    
    for t_id in lista_tarimas_id:
        story.append(Spacer(1, 1.8 * inch))
        story.append(Paragraph(f"ANEXO: TARIMA<br/><b>{t_id}</b>", style_caratula))
        story.append(Spacer(1, 0.4 * inch))
        t_bar = Table([["|||||||||||||||||||||||||||||||"], [f"*{t_id}*"]], colWidths=[540])
        t_bar.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('TEXTCOLOR', (0,0), (-1,1), colors.darkgray)]))
        story.append(t_bar)
        story.append(PageBreak())
        
        story.append(Paragraph(f"<b>DETALLE ESPECÍFICO - TARIMA {t_id}</b>", styles['Heading2']))
        sub_det = df_detalles_remision[df_detalles_remision['ID_Tarima'] == t_id]
        tabla_anexo = [[Paragraph("PO ASOCIADA", style_b), Paragraph("SKU / PRODUCTO", style_b), Paragraph("CANTIDAD", style_b)]]
        for _, s_row in sub_det.iterrows():
            tabla_anexo.append([Paragraph(str(s_row['PO']), style_t), Paragraph(str(s_row['SKU']), style_t), Paragraph(str(int(s_row['Cantidad'])), style_t)])
        t_det = Table(tabla_anexo, colWidths=[180, 180, 180])
        t_det.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F5F5F5")), ('GRID', (0,0), (-1,-1), 1, colors.white)]))
        story.append(t_det)
        story.append(PageBreak())
    if story: story.pop()
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer
# =============================================================================
# 5. RENDERIZADO DE LAS VISTAS ACTIVAS DE LA APLICACIÓN
# =============================================================================
if opcion_menu == "📊 Dashboard e Históricos":
    st.title("📊 Dashboard Planta Metales Inventario Producto")
    col_f1, col_f2 = st.columns(2)
    with col_f1: f_inicio = st.date_input("Fecha Inicial", datetime.date.today() - datetime.timedelta(days=7), key="dash_fecha_inicio_unique")
    with col_f2: f_fin = st.date_input("Fecha Final", datetime.date.today(), key="dash_fecha_final_unique")

    
    t_tar = len(st.session_state.BD_Tarimas)
    disp = len(st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible'])
    rem = len(st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Remesada'])
    
    if not st.session_state.BD_Detalle_Tarimas.empty:
        total_piezas = int(st.session_state.BD_Detalle_Tarimas['Cantidad'].sum())
        id_disp = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible']['ID_Tarima']
        piezas_disp = int(st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(id_disp)]['Cantidad'].sum())
        piezas_rem = total_piezas - piezas_disp
    else:
        total_piezas = piezas_disp = piezas_rem = 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📦 Total Tarimas", f"{t_tar} Reg.")
    m2.metric("🟢 Tarimas Disponibles", disp)
    m3.metric("🚚 Tarimas Remesadas", rem)
    m4.metric("🔢 Total Piezas", f"{total_piezas:,} PZS")
    
    st.write("---")
    st.subheader("🔍 Filtros Operativos del Proyecto")
    col_x1, col_x2, col_x3 = st.columns(3)
    df_c_filtro = st.session_state.BD_Detalle_Tarimas.copy()
    
    with col_x1:
        opciones_proy = ["Todos"] + df_c_filtro['Proyecto'].dropna().unique().tolist() if not df_c_filtro.empty and 'Proyecto' in df_c_filtro.columns else ["Todos"]
        # SE AÑADE KEY ÚNICO PARA EL PROYECTO
        proy_sel = st.selectbox("Filtrar por Proyecto Interno:", opciones_proy, key="dash_filt_proyecto_unique")
        
    with col_x2:
        if proy_sel != "Todos": df_c_filtro = df_c_filtro[df_c_filtro['Proyecto'] == proy_sel]
        opciones_po = ["Todas"] + df_c_filtro['PO'].dropna().unique().tolist() if not df_c_filtro.empty else ["Todas"]
        # SE AÑADE KEY ÚNICO PARA LA PO
        po_sel = st.selectbox("Filtrar por Orden de Compra (PO):", opciones_po, key="dash_filt_po_unique")
        
    with col_x3:
        if po_sel != "Todas": df_c_filtro = df_c_filtro[df_c_filtro['PO'] == po_sel]
        opciones_parc = ["Todas"] + df_c_filtro['Parcialidad'].dropna().unique().tolist() if not df_c_filtro.empty and 'Parcialidad' in df_c_filtro.columns else ["Todas"]
        # SE AÑADE KEY ÚNICO PARA LA PARCIALIDAD
        parc_sel = st.selectbox("Filtrar por Parcialidad:", opciones_parc, key="dash_filt_parcialidad_unique")


    st.write("---")
    st.subheader("📋 Resumen de Cumplimiento y Parcialidades")
    if not st.session_state.BD_Detalle_Tarimas.empty:
        df_completo = pd.merge(st.session_state.BD_Detalle_Tarimas, st.session_state.BD_Tarimas[['ID_Tarima', 'Estatus']], on='ID_Tarima', how='left')
        if proy_sel != "Todos": df_completo = df_completo[df_completo['Proyecto'] == proy_sel]
        if po_sel != "Todas": df_completo = df_completo[df_completo['PO'] == po_sel]
        if parc_sel != "Todas": df_completo = df_completo[df_completo['Parcialidad'] == parc_sel]
            
        if not df_completo.empty and all(col in df_completo.columns for col in ['Proyecto', 'Parcialidad', 'Descripcion']):
            resumen_avanzado = df_completo.groupby(['PO', 'Proyecto', 'Parcialidad', 'Descripcion']).agg(
                Cant_Tarimas=('ID_Tarima', 'nunique'), Total_Piezas=('Cantidad', 'sum'),
                Piezas_Disponibles=('Cantidad', lambda x: x[df_completo.loc[x.index, 'Estatus'] == 'Disponible'].sum()),
                Piezas_Remesadas=('Cantidad', lambda x: x[df_completo.loc[x.index, 'Estatus'] == 'Remesada'].sum())
            ).reset_index()
            resumen_avanzado['% Avance Salida'] = (resumen_avanzado['Piezas_Remesadas'] / resumen_avanzado['Total_Piezas'] * 100).round(1)
            resumen_avanzado['% Avance Salida'] = resumen_avanzado['% Avance Salida'].apply(lambda x: f"{x}%")
            resumen_avanzado.columns = ["Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", "Descripción del Lote", "Cant. Tarimas", "Total Piezas", "Piezas Disponibles", "Piezas Remesadas", "% Avance Salida"]
            st.dataframe(resumen_avanzado, use_container_width=True, hide_index=True)
        else: st.info("No existen registros coincidentes.")
    else: st.info("No hay datos de parcialidades cargados.")
        
    st.write("---")
    st.subheader("📋 Estado Detallado del Inventario Entarimado (Maestro)")
    if not st.session_state.BD_Tarimas.empty:
        vista_dash = st.session_state.BD_Tarimas.drop(columns=["Es_Nueva"], errors="ignore")
        st.dataframe(vista_dash, use_container_width=True)

elif opcion_menu == "🔍 Centro de Consultas":
    st.title("🔍 Centro de Consultas Avanzado")
    col_sel, col_val = st.columns(2)
    with col_sel: tipo_filtro = st.selectbox("Filtrar por:", ["Ninguno", "SKU", "PO", "Folio_Remision"])
    with col_val: valor_filtro = st.text_input("Término de búsqueda:")
    res = st.session_state.BD_Detalle_Tarimas.copy()
    if tipo_filtro == "SKU" and valor_filtro: res = res[res['SKU'].str.contains(valor_filtro, case=False)]
    elif tipo_filtro == "PO" and valor_filtro: res = res[res['PO'].str.contains(valor_filtro, case=False)]
    st.dataframe(res, use_container_width=True)

elif opcion_menu == "📦 Módulo Tarimas":
    st.title("📦 Carga de Tarimas")
    st.subheader("📋 Formato Requerido (Versión Proyectos y Parcialidades)")
    
    # 1. ENCAPSULADO DE DATOS MUESTRA DE LA PLANTILLA
    df_p = pd.DataFrame([
        {
            "Tarima": "Bulto_1", "Producto/SKU": "12-B-9016-01", "PO": "PO-10001", 
            "Proyecto": "INT-001", "Parcialidad": "Entrega_P1", 
            "Descripcion": "Reno 6", "Cantidad": 191
        }
    ])
    
    # 2. MOTOR DE DISEÑO CORPORATIVO REQUERIDO (OPENPYXL)
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    buf_p = io.BytesIO()
    with pd.ExcelWriter(buf_p, engine='openpyxl') as wr: 
        df_p.to_excel(wr, index=False, sheet_name='Plantilla_Tarimas')
        workbook, worksheet = wr.book, wr.sheets['Plantilla_Tarimas']
        
        fill_rojo = PatternFill(start_color="D32F2F", end_color="D32F2F", fill_type="solid")
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        font_data = Font(name="Calibri", size=11, bold=False, color="000000")
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        borde_t = Side(border_style="thin", color="D3D3D3")
        borde_c = Border(left=borde_t, right=borde_t, top=borde_t, bottom=borde_t)
        
        worksheet.row_dimensions[1].height = 24
        for col_idx in range(1, 8):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill, cell.font, cell.alignment = fill_rojo, font_header, align_center
            
        worksheet.row_dimensions[2].height = 20
        for col_idx in range(1, 8):
            cell = worksheet.cell(row=2, column=col_idx)
            cell.font, cell.border = font_data, borde_c
            if col_idx in [1, 3, 4, 5, 7]: cell.alignment = align_center
            else: cell.alignment = align_left
                
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            worksheet.column_dimensions[col.column_letter].width = max(max_len + 4, 15)

            
    buf_p.seek(0)
    st.download_button(label="📥 Descargar Formato de Plantilla Corporativa (.xlsx)", data=buf_p.getvalue(), file_name="plantilla_carga_tarimas_sigrama.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.write("---")
    
    # 3. RECEPCIÓN Y FILTRADO POR PASSWORD DE ARCHIVOS MAESTROS
    if not st.session_state.BD_Tarimas.empty and "Es_Nueva" not in st.session_state.BD_Tarimas.columns:
        st.session_state.BD_Tarimas["Es_Nueva"] = False

    if not is_admin: 
        st.error("🔒 Área Bloqueada: Requiere contraseña de Administrador.")
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
                if not all(col in df_ex.columns for col in columnas_requeridas):
                    st.error("❌ Error: Columnas incompatibles con el nuevo formato de proyectos.")
                else:
                    if not st.session_state.BD_Tarimas.empty: st.session_state.BD_Tarimas["Es_Nueva"] = False
                    for t_orig in df_ex['Tarima'].unique():
                        nuevo_id_tpm = f"TPM-{(len(st.session_state.BD_Tarimas) + 1):04d}"
                        n_t = {"ID_Tarima": nuevo_id_tpm, "Tarima_Origen_Excel": t_orig, "Fecha_Creacion": datetime.datetime.now().strftime("%d/%m/%Y"), "Ubicacion_Actual": "Metales", "Creado_Por": oper, "Tipo_Tarima": tipo_t, "Estatus": "Disponible", "Es_Nueva": True}
                        st.session_state.BD_Tarimas = pd.concat([st.session_state.BD_Tarimas, pd.DataFrame([n_t])], ignore_index=True)
                        
                        items = df_ex[df_ex['Tarima'] == t_orig]
                        for _, item in items.iterrows():
                            st.session_state.BD_Detalle_Tarimas = pd.concat([st.session_state.BD_Detalle_Tarimas, pd.DataFrame([{
                                "ID_Detalle": len(st.session_state.BD_Detalle_Tarimas) + 1, "ID_Tarima": nuevo_id_tpm, 
                                "SKU": item['Producto/SKU'], "PO": item['PO'], "Proyecto": item['Proyecto'],
                                "Parcialidad": item['Parcialidad'], "Descripcion": item['Descripcion'], "Cantidad": item['Cantidad']
                            }])], ignore_index=True)
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                    st.success("¡Inventario guardado y respaldado con éxito!")
                    st.rerun()
            except Exception as e: st.error(f"Error: {e}")
            
    # 4. CUADRÍCULA INTERACTIVA DE IMPRESIÓN SELECCIONADA EN MEMORIA
    if not st.session_state.BD_Tarimas.empty:
        st.write("---")
        st.subheader("🖨️ Panel de Impresión Masiva de Tarimas")
        df_estilado = st.session_state.BD_Tarimas.style.apply(lambda r: ['background-color: #FFF59D' if r['Es_Nueva'] else '' for _ in r], axis=1)
        seleccion_tabla = st.dataframe(df_estilado, use_container_width=True, column_order=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus"], on_select="rerun", selection_mode="multi-row")
        filas_seleccionadas = seleccion_tabla.get("selection", {}).get("rows", [])
        
        if filas_seleccionadas:
            tarimas_elegidas = st.session_state.BD_Tarimas.iloc[filas_seleccionadas]['ID_Tarima'].tolist()
            if len(tarimas_elegidas) == 1:
                t_imp = tarimas_elegidas[0]
                st.download_button(label=f"📥 Descargar PDF Tarima #{t_imp}", data=generar_pdf_tarima(t_imp), file_name=f"Tarima_{t_imp}.pdf", mime="application/pdf")
            else:
                if st.button("📦 Unificar y Preparar Lote de Impresión"):
                    buf_l = io.BytesIO()
                    doc_l = SimpleDocTemplate(buf_l, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
                    story_l, styles = [], getSampleStyleSheet()
                    for t_imp in tarimas_elegidas:
                        det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == t_imp]
                        t_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == t_imp].iloc[0]
                        story_l.append(Spacer(1, 1.8 * inch))
                        story_l.append(Paragraph(f"TARIMA<br/><br/><b>#{t_imp}</b>", ParagraphStyle('G_L', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1)))
                        story_l.append(PageBreak())
                        story_l.append(Paragraph(f"<b>Detalle Interno - Tarima #{t_imp}</b>", styles['Heading2']))
                        story_l.append(Paragraph(f"<b>Operador:</b> {t_info['Creado_Por']} | <b>Fecha:</b> {t_info['Fecha_Creacion']}", styles['Normal']))
                        story_l.append(Spacer(1, 0.3 * inch))
                        for _, item in det.iterrows():
                            art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
                            nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
                            story_l.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", styles['Normal']))
                            story_l.append(Spacer(1, 0.4 * inch))
                            story_l.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", ParagraphStyle('NG_L', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)))
                        story_l.append(PageBreak())
                    if story_l: story_l.pop()
                    doc_l.build(story_l, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
                    st.download_button(label="📥 Descargar Lote Completo (PDF)", data=buf_l.getvalue(), file_name="Lote_Tarimas.pdf", mime="application/pdf")
        else:
            st.warning("Seleccione una o más filas en la tabla usando las casillas de verificación para activar la descarga.")




# =============================================================================
# SECCIÓN 14: PANEL DE MANTENIMIENTO, EDICIÓN EN CALIENTE Y PURGA DE REGISTROS
# =============================================================================
elif opcion_menu == "⚙️ Mantenimiento y Catálogos":
    st.title("⚙️ Panel de Mantenimiento Avanzado del Sistema")
    st.warning("⚠️ Acción Crítica: Las modificaciones realizadas impactan directamente en GitHub.")
    
    if "BD_Lideres" not in st.session_state:
        st.session_state.BD_Lideres = pd.DataFrame([
            {"ID_Lider": "LID-01", "Nombre_Lider": "Jesus Morales", "Area": "Metales", "Estatus": "Activo"}
        ])

    tab1, tab2, tab3 = st.tabs(["📝 Ajustar Cantidades", "👤 Catálogo de Líderes", "🚨 Purga de Datos"])

    # --- PESTAÑA 1: AJUSTAR CANTIDADES EN CALIENTE ---
    with tab1:
        st.subheader("✏️ Edición Rápida de Inventario (Detalle Tarimas)")
        if not st.session_state.BD_Detalle_Tarimas.empty:
            df_editable = st.data_editor(
                st.session_state.BD_Detalle_Tarimas, 
                use_container_width=True, 
                disabled=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Proyecto", "Parcialidad", "Descripcion"], 
                hide_index=True
            )
            if st.button("💾 Guardar Cambios de Cantidades en GitHub"):
                st.session_state.BD_Detalle_Tarimas = df_editable
                subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                st.success("✅ Cantidades corregidas y sincronizadas con éxito."); st.rerun()
        else: 
            st.info("No hay registros en el detalle de tarimas para modificar.")

    # --- PESTAÑA 2: ADMISTRACIÓN DE PERSONAL DE LÍDERES ---
    with tab2:
        st.subheader("👤 Administración del Personal de Líderes")
        with st.expander("➕ Dar de Alta Nuevo Líder"):
            c_l1, c_l2 = st.columns(2)
            with c_l1: nuevo_nom = st.text_input("Nombre Completo:")
            with c_l2: nueva_area = st.text_input("Área:", "Metales")
            if st.button("➕ Registrar Líder"):
                if nuevo_nom:
                    n_row = {"ID_Lider": f"LID-{(len(st.session_state.BD_Lideres) + 1):02d}", "Nombre_Lider": nuevo_nom, "Area": nueva_area, "Estatus": "Activo"}
                    st.session_state.BD_Lideres = pd.concat([st.session_state.BD_Lideres, pd.DataFrame([n_row])], ignore_index=True)
                    subir_excel_a_github("BD_Lideres.xlsx", st.session_state.BD_Lideres)
                    st.success("Líder registrado."); st.rerun()
        df_l_edit = st.data_editor(st.session_state.BD_Lideres, use_container_width=True, hide_index=True)
        if st.button("💾 Sincronizar Cambios de Líderes"):
            st.session_state.BD_Lideres = df_l_edit
            subir_excel_a_github("BD_Lideres.xlsx", st.session_state.BD_Lideres)
            st.success("Catálogo de líderes sincronizado.")

    # --- PESTAÑA 3: PURGA DETALLADA Y MASIVA DE REGISTROS ---
    with tab3:
        st.subheader("🚨 Reset de Fábrica y Purga de Datos Controlada")
        metodo_purga = st.radio("Método de Purga:", ["❌ Purga Total Automática (Reset Completo)", "🔍 Seleccionar Registros Específicos para Eliminar"], horizontal=True)
        st.write("---")
        
        if metodo_purga == "❌ Purga Total Automática (Reset Completo)":
            st.error("⚠️ Peligro: Esta acción vaciará por completo todos los registros históricos en GitHub a ceros.")
            if st.checkbox("Confirmo que deseo aplicar un Reset de Fábrica Total.", key="chk_total"):
                if st.button("🗑️ EJECUTAR PURGA MAESTRA TOTAL"):
                    st.session_state.BD_Tarimas = pd.DataFrame(columns=st.session_state.BD_Tarimas.columns)
                    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=st.session_state.BD_Detalle_Tarimas.columns)
                    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=st.session_state.BD_Datos_Generales_Remision.columns)
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success("💥 Sistema purgado por completo a ceros de forma masiva."); st.rerun()
        else:
            st.info("💡 Consejo: Seleccione una o más casillas en las tablas inferiores para habilitar los botones de remoción parcial.")
            st.markdown("### 📦 1. Eliminar Tarimas del Inventario")
            if not st.session_state.BD_Tarimas.empty:
                df_tar_vista = st.session_state.BD_Tarimas.copy().drop(columns=["Es_Nueva"], errors="ignore")
                sel_tarimas = st.dataframe(df_tar_vista, use_container_width=True, on_select="rerun", selection_mode="multi-row", key="tabla_purga_tarimas")
                filas_tar = sel_tarimas.get("selection", {}).get("rows", [])
                if filas_tar:
                    ids_tar_eliminar = st.session_state.BD_Tarimas.iloc[filas_tar]['ID_Tarima'].tolist()
                    st.warning(f"⚠️ Seleccionadas para eliminar: {', '.join(ids_tar_eliminar)}")
                    if st.button("🗑️ Eliminar Tarimas Seleccionadas"):
                        st.session_state.BD_Tarimas = st.session_state.BD_Tarimas[~st.session_state.BD_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                        st.session_state.BD_Detalle_Tarimas = st.session_state.BD_Detalle_Tarimas[~st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                        subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                        subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                        st.success("✅ Registros de Tarimas removidos e inyectados en GitHub."); st.rerun()
            else: st.write("No hay tarimas registradas en el inventario activo.")
                
            st.write("---")
            st.markdown("### 🚚 2. Eliminar Remisiones de Salida")
            if not st.session_state.BD_Datos_Generales_Remision.empty:
                sel_remisiones = st.dataframe(st.session_state.BD_Datos_Generales_Remision, use_container_width=True, on_select="rerun", selection_mode="multi-row", key="tabla_purga_remisiones")
                filas_rem = sel_remisiones.get("selection", {}).get("rows", [])
                if filas_rem:
                    ids_rem_eliminar = st.session_state.BD_Datos_Generales_Remision.iloc[filas_rem]['Folio_Remision'].tolist()
                    tarimas_afectadas = []
                    for idx in filas_rem: tarimas_afectadas.extend(st.session_state.BD_Datos_Generales_Remision.iloc[idx]['Tarimas_Asociadas'])
                    st.warning(f"⚠️ Seleccionadas para eliminar: {', '.join(ids_rem_eliminar)}")
                    if st.button("🗑️ Eliminar Remisiones Seleccionadas"):
                        st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(tarimas_afectadas), 'Estatus'] = 'Disponible'
                        st.session_state.BD_Datos_Generales_Remision = st.session_state.BD_Datos_Generales_Remision[~st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].isin(ids_rem_eliminar)]
                        subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                        subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                        st.success("✅ Folios de remisión eliminados y tarimas restauradas a Disponible."); st.rerun()
            else: st.write("No hay folios de remisiones emitidos en el historial.")
