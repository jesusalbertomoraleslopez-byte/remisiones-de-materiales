import streamlit as st
import pandas as pd
import datetime
import io
import base64
import requests
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# 1. INICIALIZACIÓN DE LA PLATAFORMA WEB
st.set_page_config(page_title="Remisiones de Materiales", layout="wide", page_icon="📦")

# 2. CARGA DEL BANNER RESPONSIVO CORPORATIVO
try:
    banner_img = Image.open("REMISIONES APP.png")
    st.image(banner_img, use_container_width=True)
except FileNotFoundError:
    st.warning("⚠️ Cargando interfaz gráfica del banner superior...")
st.write("")
# 3. BASE DE DATOS EN MEMORIA (PANDAS + CACHÉ DE SESIÓN)
if "BD_Articulos" not in st.session_state:
    st.session_state.BD_Articulos = pd.DataFrame([
        {"SKU": "12-B-9016-01", "Nombre": "Lámina Galvanizada Sigrama", "Calibre_Espesor": "Calibre 22", "Dimensiones_Pieza": "3x10 ft", "Acabado_Superficial": "Zintro"},
        {"SKU": "SKU-002", "Nombre": "Placa de Acero Comercial", "Calibre_Espesor": "1/4 pulgada", "Dimensiones_Pieza": "4x8 ft", "Acabado_Superficial": "Negro"}
    ])

if "BD_Tarimas" not in st.session_state:
    st.session_state.BD_Tarimas = pd.DataFrame(columns=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus", "Es_Nueva"])

if "BD_Detalle_Tarimas" not in st.session_state:
    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Proyecto", "Parcialidad", "Descripcion", "Cantidad"])

if "BD_Datos_Generales_Remision" not in st.session_state:
    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=["ID_Remision", "Folio_Remision", "Fecha_Hora_Salida", "Nombre_Emisor", "Direccion_Emisor", "Nombre_Receptor", "Direccion_Receptor", "Tarimas_Asociadas"])

if "BD_Lideres" not in st.session_state:
    st.session_state.BD_Lideres = pd.DataFrame([{"ID_Lider": "LID-01", "Nombre_Lider": "Jesus Morales", "Area": "Metales", "Estatus": "Activo"}])
# 4. CAPA DE SEGURIDAD MULTINIVEL
st.sidebar.title("🔐 Control de Acceso")
admin_pass_input = st.sidebar.text_input("Contraseña Administrador:", type="password", key="sec_admin_pass")

def es_admin():
    try: return admin_pass_input == st.secrets["admin_password"]
    except KeyError: return False

is_admin = es_admin()
if is_admin: st.sidebar.success("Modo Administrador Activo")
else: st.sidebar.warning("Modo Consulta Activo")

st.sidebar.write("---")
st.sidebar.title("🛠️ Área de Soporte")
super_pass_input = st.sidebar.text_input("Contraseña de Soporte / IT:", type="password", key="sec_super_pass")
is_super = (super_pass_input == "SigramaMetales2025")
if is_super: st.sidebar.success("⚡ Modo Superusuario Activo")

st.sidebar.title("🧭 Navegación")
lista_modulos = ["📊 Dashboard e Históricos", "🔍 Centro de Consultas", "📦 Módulo Tarimas", "🚚 Módulo Remisiones"]
if is_super: lista_modulos.append("⚙️ Mantenimiento y Catálogos")
opcion_menu = st.sidebar.radio("Seleccione un Módulo:", lista_modulos, key="sec_nav_radio")
# 5. MOTOR DE PERSISTENCIA AUTOMÁTICO EN GITHUB VIA API
def subir_excel_a_github(file_name, dataframe_to_save):
    try:
        GITHUB_TOKEN = st.secrets["github_token"]
        REPO_OWNER = "jesusalbertomoraleslopez-byte"
        REPO_NAME = "remisiones-de-materiales"
        BRANCH = "main"
        
        buffer_git = io.BytesIO()
        with pd.ExcelWriter(buffer_git, engine='openpyxl') as writer:
            dataframe_to_save.to_excel(writer, index=False, sheet_name='Datos')
        
        base64_content = base64.b64encode(buffer_git.getvalue()).decode("utf-8")
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{file_name}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
        
        res_get = requests.get(url, headers=headers)
        sha = res_get.json().get("sha") if res_get.status_code == 200 else None
        
        payload = {"message": f"Sync: {file_name}", "content": base64_content, "branch": BRANCH}
        if sha: payload["sha"] = sha
            
        res_put = requests.put(url, json=payload, headers=headers)
        if res_put.status_code in [200, 201]:
            st.sidebar.success(f"💾 Guardado en GitHub: {file_name}")
            return True
        return False
    except Exception: return False
# 6. CAPA GRÁFICA FIJA DEL PDF CORPORATIVO
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
    texto_legal = "PROHIBIDA LA REPRODUCCIÓN TOTAL O PARCIAL SIN AUTORIZACIÓN DE INDUSTRIA SIGRAMA S.A. DE C.V."
    canvas.drawString(95, 37, texto_legal)
    canvas.restoreState()
# 7. GENERADOR DOCUMENTAL DE TARIMA INDIVIDUAL
def generar_pdf_tarima(id_tarima):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    tarima_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == id_tarima].iloc[0]
    detalles = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == id_tarima]
    
    style_g = ParagraphStyle('G', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1)
    story.append(Spacer(1, 1.8 * inch))
    story.append(Paragraph(f"TARIMA<br/><br/><b>#{id_tarima}</b>", style_g))
    story.append(PageBreak())
    
    style_n = styles['Normal']
    style_ng = ParagraphStyle('NG', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)
    story.append(Paragraph(f"<b>Detalle Interno - Tarima #{id_tarima}</b>", styles['Heading2']))
    story.append(Paragraph(f"<b>Operador:</b> {tarima_info['Creado_Por']} | <b>Fecha:</b> {tarima_info['Fecha_Creacion']}", style_n))
    story.append(Spacer(1, 0.3 * inch))
    
    for _, item in detalles.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
        story.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", style_n))
        story.append(Spacer(1, 0.4 * inch))
        story.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", style_ng))
        
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer.getvalue()
# 8. GENERADOR DOCUMENTAL DE LA REMISIÓN GENERAL
def generar_pdf_remision_general(datos_remision, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    style_blanco_bold = ParagraphStyle('BB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_normal_bold = ParagraphStyle('NB', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=9)
    style_normal_text = ParagraphStyle('NT', parent=styles['Normal'], fontSize=9)
    
    story.append(Spacer(1, 0.1 * inch))
    t_header_embarque = Table([[Paragraph("LISTADO DE EMBARQUE", style_blanco_bold)]], colWidths=[7.5 * inch])
    t_header_embarque.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header_embarque)
    
    datos_panel = [
        [Paragraph("FOLIO:", style_normal_bold), Paragraph(str(datos_remision['Folio_Remision']), style_normal_text), Paragraph("FECHA:", style_normal_bold), Paragraph(str(datos_remision['Fecha_Hora_Salida']), style_normal_text)],
        [Paragraph("LÍDER:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Emisor']), style_normal_text), Paragraph("NO. ALMACÉN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Emisor']), style_normal_text)],
        [Paragraph("DESTINO:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Receptor']), style_normal_text), "", ""],
        [Paragraph("DIRECCIÓN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Receptor']), style_normal_text), "", ""]
    ]
    t_panel = Table(datos_panel, colWidths=[1.5 * inch, 2.25 * inch, 1.5 * inch, 2.25 * inch])
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
        
    t_mat = Table(tabla_materiales, colWidths=[1.0 * inch, 1.0 * inch, 1.5 * inch, 2.5 * inch, 1.5 * inch])
    t_mat.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575"))]))
    story.append(t_mat)
    story.append(Spacer(1, 0.3 * inch))
    
    t_firmas = Table([[Paragraph("ENTREGA", style_normal_bold), Paragraph("RECIBE", style_normal_bold)]], colWidths=[3.75 * inch, 3.75 * inch])
    t_firmas.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('LINEBELOW', (0,0), (0,0), 1, colors.black), ('LINEBELOW', (1,0), (1,0), 1, colors.black), ('TOPPADDING', (0,0), (-1,-1), 25)]))
    story.append(t_firmas)
    
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer.getvalue()
# 9. GENERADOR DOCUMENTAL DE ANEXOS DE TARIMA CON INTERLINEADO CORRECTO
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
        
        t_bar = Table([["|||||||||||||||||||||||||||||||"], [f"*{t_id}*"]], colWidths=[6.5 * inch])
        t_bar.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('TEXTCOLOR', (0,0), (-1,1), colors.darkgray)]))
        story.append(t_bar)
        story.append(PageBreak())
        
        story.append(Paragraph(f"<b>DETALLE ESPECÍFICO - TARIMA {t_id}</b>", styles['Heading2']))
        story.append(Spacer(1, 0.1 * inch))
        
        sub_det = df_detalles_remision[df_detalles_remision['ID_Tarima'] == t_id]
        tabla_anexo = [[Paragraph("PO ASOCIADA", style_b), Paragraph("SKU / PRODUCTO", style_b), Paragraph("CANTIDAD", style_b)]]
        for _, s_row in sub_det.iterrows():
            tabla_anexo.append([Paragraph(str(s_row['PO']), style_t), Paragraph(str(s_row['SKU']), style_t), Paragraph(str(int(s_row['Cantidad'])), style_t)])
            
        t_det = Table(tabla_anexo, colWidths=[2.5 * inch, 2.5 * inch, 2.5 * inch])
        t_det.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F5F5F5")), ('GRID', (0,0), (-1,-1), 1, colors.white)]))
        story.append(t_det)
        story.append(PageBreak())
        
    if story: story.pop()
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer.getvalue()
# 10. MÓDULO INFORMATIVO: DASHBOARD GENERAL
if opcion_menu == "📊 Dashboard e Históricos":
    st.title("📊 Dashboard Planta Metales Inventario Producto")
    col_f1, col_f2 = st.columns(2)
    with col_f1: f_inicio = st.date_input("Fecha Inicial", datetime.date.today() - datetime.timedelta(days=7), key="dash_f_inicio_u")
    with col_f2: f_fin = st.date_input("Fecha Final", datetime.date.today(), key="dash_f_fin_u")
    
    t_tar = len(st.session_state.BD_Tarimas)
    disp = len(st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible'])
    rem = len(st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Remesada'])
    
    if not st.session_state.BD_Detalle_Tarimas.empty:
        total_piezas = int(st.session_state.BD_Detalle_Tarimas['Cantidad'].sum())
        id_disp = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible']['ID_Tarima']
        piezas_disp = int(st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(id_disp)]['Cantidad'].sum())
        piezas_rem = total_piezas - piezas_disp
    else: total_piezas = piezas_disp = piezas_rem = 0

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
        proy_sel = st.selectbox("Filtrar por Proyecto Interno:", opciones_proy, key="dash_f_proy_u")
    with col_x2:
        if proy_sel != "Todos": df_c_filtro = df_c_filtro[df_c_filtro['Proyecto'] == proy_sel]
        opciones_po = ["Todas"] + df_c_filtro['PO'].dropna().unique().tolist() if not df_c_filtro.empty else ["Todas"]
        po_sel = st.selectbox("Filtrar por Orden de Compra (PO):", opciones_po, key="dash_f_po_u")
    with col_x3:
        if po_sel != "Todas": df_c_filtro = df_c_filtro[df_c_filtro['PO'] == po_sel]
        opciones_parc = ["Todas"] + df_c_filtro['Parcialidad'].dropna().unique().tolist() if not df_c_filtro.empty and 'Parcialidad' in df_c_filtro.columns else ["Todas"]
        parc_sel = st.selectbox("Filtrar por Parcialidad:", opciones_parc, key="dash_f_parc_u")
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
                       # LÍNEA CORREGIDA CON EL NUEVO NOMBRE DE ENCABEZADO DE COLUMNA
            resumen_avanzado.columns = [
                "Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", 
                "Descripción de Proyecto Planta Rio", "Cant. Tarimas", "Total Piezas", 
                "Piezas Disponibles", "Piezas Remesadas", "% Avance Salida"
            ]
            st.dataframe(resumen_avanzado, use_container_width=True, hide_index=True)
        else: st.info("No existen registros coincidentes con el filtro.")
    else: st.info("No hay datos de parcialidades.")
    st.write("---")
    st.subheader("📋 Estado Detallado del Inventario Entarimado (Maestro)")
    if not st.session_state.BD_Tarimas.empty:
        st.dataframe(st.session_state.BD_Tarimas.drop(columns=["Es_Nueva"], errors="ignore"), use_container_width=True)
    else: st.info("No hay tarimas registradas.")
# 12. MÓDULO DE BÚSQUEDA EXCEL Y CONSULTAS
elif opcion_menu == "🔍 Centro de Consultas":
    st.title("🔍 Centro de Consultas Avanzado")
    col_sel, col_val = st.columns(2)
    with col_sel: tipo_filtro = st.selectbox("Filtrar por:", ["Ninguno", "SKU", "PO", "Folio_Remision"], key="consult_tipo_u")
    with col_val: valor_filtro = st.text_input("Término de búsqueda:", key="consult_val_u")
    res = st.session_state.BD_Detalle_Tarimas.copy()
    if tipo_filtro == "SKU" and valor_filtro: res = res[res['SKU'].str.contains(valor_filtro, case=False)]
    elif tipo_filtro == "PO" and valor_filtro: res = res[res['PO'].str.contains(valor_filtro, case=False)]
    elif tipo_filtro == "Folio_Remision" and valor_filtro:
        rem_m = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].str.contains(valor_filtro, case=False)]
        tar_v = [t for r in rem_m['Tarimas_Asociadas'] for t in r]
        res = res[res['ID_Tarima'].isin(tar_v)]
    st.dataframe(res, use_container_width=True)
# =============================================================================
# SECCIÓN 13: MÓDULO TARIMAS - FORMULACIÓN DE PLANTILLA ESTILIZADA SIGRAMA
# =============================================================================
elif opcion_menu == "📦 Módulo Tarimas":
    st.title("📦 Carga de Tarimas")
    st.subheader("📋 Formato Requerido (Versión Proyectos y Parcialidades)")
    
    # DataFrame muestra con la estructura avanzada requerida
    df_p = pd.DataFrame([
        {
            "Tarima": "Bulto_1", "Producto/SKU": "12-B-9016-01", "PO": "PO-10001", 
            "Proyecto": "INT-001", "Parcialidad": "Entrega_P1", 
            "Descripcion": "Reno 6", "Cantidad": 191
        }
    ])
    
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    buf_p = io.BytesIO()
    with pd.ExcelWriter(buf_p, engine='openpyxl') as wr: 
        df_p.to_excel(wr, index=False, sheet_name='Plantilla_Tarimas')
        workbook, worksheet = wr.book, wr.sheets['Plantilla_Tarimas']
        
        # Estilos corporativos de Industria Sigrama
        fill_rojo = PatternFill(start_color="D32F2F", end_color="D32F2F", fill_type="solid")
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        font_data = Font(name="Calibri", size=11, bold=False, color="000000")
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        borde_t = Side(border_style="thin", color="D3D3D3")
        borde_c = Border(left=borde_t, right=borde_t, top=borde_t, bottom=borde_t)
        
        # Formato de Encabezados (Fila 1)
        worksheet.row_dimensions[1].height = 24
        for col_idx in range(1, 8):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill, cell.font, cell.alignment = fill_rojo, font_header, align_center
            
        # Formato de Celdas de Datos (Fila 2)
        worksheet.row_dimensions[2].height = 20
        for col_idx in range(1, 8):
            cell = worksheet.cell(row=2, column=col_idx)
            cell.font, cell.border = font_data, borde_c
            if col_idx in [1, 3, 4, 5, 7]: 
                cell.alignment = align_center
            else: 
                cell.alignment = align_left
                
        # Auto-ajuste de ancho de columnas sin errores de atributo
        # Importación explícita del módulo de utilerías indexado de openpyxl
        from openpyxl.utils import get_column_letter
        
        # Bucle corregido para iterar de manera segura extrayendo el índice de la primera celda
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)  # <-- CAMBIO CLAVE DE COMPATIBILIDAD
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 15)

            
    buf_p.seek(0)
    st.download_button(label="📥 Descargar Formato de Plantilla Corporativa (.xlsx)", data=buf_p.getvalue(), file_name="plantilla_carga_tarimas_sigrama.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.write("---")
# =============================================================================
# SECCIÓN 14: MÓDULO TARIMAS - CARGA MASIVA E INYECCIÓN RELACIONAL A GITHUB
# =============================================================================
    if not st.session_state.BD_Tarimas.empty and "Es_Nueva" not in st.session_state.BD_Tarimas.columns:
        st.session_state.BD_Tarimas["Es_Nueva"] = False

    if not is_admin: 
        st.error("🔒 Área Bloqueada: Se requiere contraseña de Administrador para subir plantillas.")
    else:
        st.success("🔓 Acceso Autorizado.")
        arch = st.file_uploader("Suba el Excel con Formato de Proyectos", type=["xlsx"])
        col_t1, col_t2 = st.columns(2)
        with col_t1: tipo_t = st.selectbox("Tipo de Tarima (Embalaje):", ["Cuadrada", "Rectangular"])
        with col_t2: oper = st.text_input("Líder Responsable de Captura:", "Jesus Morales")
        
        if arch and st.button("Procesar e Integrar Plantilla Avanzada"):
            try:
                df_ex = pd.read_excel(arch)
                columnas_requeridas = ["Tarima", "Producto/SKU", "PO", "Proyecto", "Parcialidad", "Descripcion", "Cantidad"]
                if not all(col in df_ex.columns for col in columnas_requeridas):
                    st.error("❌ Error: El archivo no coincide con las nuevas columnas requeridas.")
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
                    st.success("¡Inventario sincronizado y respaldado con éxito en GitHub!")
                    st.rerun()
            except Exception as e: st.error(f"Error en conector: {e}")
# =============================================================================
# SECCIÓN 15: MÓDULO TARIMAS - PANEL DE IMPRESIÓN SELECCIONADA POR LOTE
# =============================================================================
    if not st.session_state.BD_Tarimas.empty:
        st.write("---")
        st.subheader("🖨️ Panel de Impresión Masiva de Tarimas")
        df_estilado = st.session_state.BD_Tarimas.style.apply(lambda r: ['background-color: #FFF59D' if r['Es_Nueva'] else '' for _ in r], axis=1)
        seleccion_tabla = st.dataframe(df_estilado, use_container_width=True, column_order=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus"], on_select="rerun", selection_mode="multi-row", key="tabla_master_impresion_tarimas_u")
        filas_seleccionadas = seleccion_tabla.get("selection", {}).get("rows", [])
        
        if filas_seleccionadas:
            tarimas_elegidas = st.session_state.BD_Tarimas.iloc[filas_seleccionadas]['ID_Tarima'].tolist()
            if len(tarimas_elegidas) == 1:
                t_imp = tarimas_elegidas[0]
                st.download_button(label=f"📥 Descargar PDF Tarima #{t_imp}", data=generar_pdf_tarima(t_imp), file_name=f"Tarima_{t_imp}.pdf", mime="application/pdf", key="btn_individual_tar_dl_u")
            else:
                if st.button("📦 Unificar y Preparar Lote de Impresión Continua"):
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
                    st.download_button(label="📥 Descargar Lote Completo Unificado (PDF)", data=buf_l.getvalue(), file_name="Lote_Tarimas.pdf", mime="application/pdf", key="btn_lote_masivo_dl_u")
        else:
            st.warning("Seleccione una o más filas en la tabla usando las casillas de verificación de la izquierda.")
# =============================================================================
# SECCIÓN 16: MÓDULO REMISIONES - CONTROL DE DESPACHOS Y ASIGNACIÓN DE LÍDERES
# =============================================================================
elif opcion_menu == "🚚 Módulo Remisiones":
    st.title("🚚 Generación de Remisiones de Salida")
    t_disp = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible']['ID_Tarima'].tolist()
    if not t_disp: st.warning("⚠️ No existen tarimas disponibles bajo estatus 'Disponible' en patio.")
    else:
        t_sel = st.multiselect("Seleccione Tarimas para este Envío:", options=t_disp, key="ms_remisiones_active_u")
        col_e, col_r = st.columns(2)
        with col_e:
            if "BD_Lideres" in st.session_state and not st.session_state.BD_Lideres.empty:
                lista_nombres_lideres = st.session_state.BD_Lideres['Nombre_Lider'].unique().tolist()
                nom_e = st.selectbox("Líder / Emisor Autorizado:", options=lista_nombres_lideres, key="rem_lider_sel_unique")
            else:
                nom_e = st.selectbox("Líder / Emisor Autorizado:", options=["Jesus Morales", "Supervisor General"], key="rem_lider_backup_unique")
            dir_e = st.text_input("Almacén de Origen:", "Metales", key="ti_origen_rem_u")
        with col_r:
            nom_r = st.text_input("Receptor / Cliente:", "Galvatec Industrias", key="ti_cliente_rem_u")
            dir_r = st.text_input("Dirección Destino:", "Prol. Valle Guadiana 919, Parque Industrial II, 35078 Gómez Palacio, Dgo.", key="ti_destino_rem_u")
        if not is_admin: st.error("🔒 Operación Bloqueada: Se requiere contraseña de Administrador en la barra lateral.")
        else:
            if st.button("🚀 Confirmar Salida y Generar Nueva Remisión"):
                if not t_sel or not nom_r: st.error("Complete todos los campos obligatorios.")
                else:
                    fol = f"E00{27 + len(st.session_state.BD_Datos_Generales_Remision)}"
                    reg = {"ID_Remision": len(st.session_state.BD_Datos_Generales_Remision)+1, "Folio_Remision": fol, "Fecha_Hora_Salida": datetime.datetime.now().strftime("%d/%m/%Y"), "Nombre_Emisor": nom_e, "Direccion_Emisor": dir_e, "Nombre_Receptor": nom_r, "Direccion_Receptor": dir_r, "Tarimas_Asociadas": t_sel}
                    st.session_state.BD_Datos_Generales_Remision = pd.concat([st.session_state.BD_Datos_Generales_Remision, pd.DataFrame([reg])], ignore_index=True)
                    st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(t_sel), 'Estatus'] = 'Remesada'
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success(f"✅ ¡Remisión {fol} Generada y Respaldada de forma permanente!"); st.rerun()
                    
    if not st.session_state.BD_Datos_Generales_Remision.empty:
        st.write("---")
        st.subheader("🖨️ Descarga Documental de Remisiones")
        r_sel = st.selectbox("Seleccione Folio para Descarga:", st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].unique(), key="rem_download_folio_sel")
        row = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'] == r_sel].iloc[0]
        df_det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(row['Tarimas_Asociadas'])]
        c1, c2 = st.columns(2)
        with c1: st.download_button("📥 Descargar Remisión (PDF)", data=generar_pdf_remision_general(row, df_det), file_name=f"Remision_{r_sel}.pdf", key="btn_dl_rem_pdf", mime="application/pdf")
        with c2: st.download_button("📥 Descargar Anexo Tarimas (PDF)", data=generar_pdf_anexo_tarimas(row['Tarimas_Asociadas'], df_det), file_name=f"Anexo_{r_sel}.pdf", key="btn_dl_anexo_pdf", mime="application/pdf")
# =============================================================================
# SECCIÓN 17: PANEL DE MANTENIMIENTO, EDICIÓN EN CALIENTE Y PURGA (SUPERUSER)
# =============================================================================
elif opcion_menu == "⚙️ Mantenimiento y Catálogos":
    st.title("⚙️ Panel de Mantenimiento Avanzado del Sistema")
    st.warning("⚠️ Acción Crítica: Las modificaciones realizadas impactan directamente en GitHub.")
    if "BD_Lideres" not in st.session_state:
        st.session_state.BD_Lideres = pd.DataFrame([{"ID_Lider": "LID-01", "Nombre_Lider": "Jesus Morales", "Area": "Metales", "Estatus": "Activo"}])

    tab1, tab2, tab3 = st.tabs(["📝 Ajustar Cantidades", "👤 Catálogo de Líderes", "🚨 Purga de Datos"])

    with tab1:
        st.subheader("✏️ Edición Rápida de Inventario")
        if not st.session_state.BD_Detalle_Tarimas.empty:
            df_editable = st.data_editor(st.session_state.BD_Detalle_Tarimas, use_container_width=True, disabled=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Proyecto", "Parcialidad", "Descripcion"], hide_index=True, key="editor_mantenimiento_cantidades")
            if st.button("💾 Guardar Cambios de Cantidades en GitHub"):
                st.session_state.BD_Detalle_Tarimas = df_editable
                subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                st.success("✅ Cantidades corregidas."); st.rerun()
        else: st.info("No hay registros para modificar.")

    # --- PESTAÑA 2: ADMINISTRACIÓN DE LÍDERES CON CARGA MASIVA EXCEL ---
    # --- PESTAÑA 2: ADMINISTRACIÓN DE LÍDERES CON CARGA MASIVA EXCEL (CORREGIDO) ---
    with tab2:
        st.subheader("👤 Administración del Personal de Líderes")
        
        # 1. GENERACIÓN Y DESCARGA DE PLANTILLA DE PERSONAL
        df_l_plantilla = pd.DataFrame([
            {"Nombre_Lider": "Nombre Ejemplo 1", "Area": "Metales"},
            {"Nombre_Lider": "Nombre Ejemplo 2", "Area": "Embarques"}
        ])
        
        # Importaciones requeridas explícitas dentro del bloque de control
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        buf_l = io.BytesIO()
        with pd.ExcelWriter(buf_l, engine='openpyxl') as wr_l:
            df_l_plantilla.to_excel(wr_l, index=False, sheet_name='Plantilla_Lideres')
            
            # Obtener hoja activa
            ws_l = wr_l.sheets['Plantilla_Lideres']
            
            # CORRECCIÓN DE SINTAXIS AUTOMÁTICA (Se eliminó el prefijo 'openpyxl.')
            fill_header_lider = PatternFill(start_color="757575", end_color="757575", fill_type="solid")
            font_header_lider = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            
            for col_idx in range(1, 3):
                cell = ws_l.cell(row=1, column=col_idx)
                cell.font = font_header_lider
                cell.fill = fill_header_lider
        
        st.write("Descarga el formato base para rellenar la lista de personal en Excel:")
        st.download_button(
            label="📥 Descargar Plantilla de Personal (.xlsx)",
            data=buf_l.getvalue(),
            file_name="plantilla_lideres_sigrama.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_download_plantilla_lideres_u"
        )

        
        st.write("---")
        
        # 2. CARGADOR DE ARCHIVO EXCEL DE LÍDERES (INTEGRACIÓN MASIVA)
        arch_lideres = st.file_uploader("Suba la Plantilla de Personal Rellenada:", type=["xlsx"], key="uploader_lideres_masivo_u")
        
        if arch_lideres and st.button("🚀 Procesar e Integrar Personal Masivo"):
            try:
                df_l_excel = pd.read_excel(arch_lideres)
                columnas_l_req = ["Nombre_Lider", "Area"]
                
                if not all(col in df_l_excel.columns for col in columnas_l_req):
                    st.error("❌ Error: Las columnas del archivo no coinciden. Use el formato: Nombre_Lider, Area")
                else:
                    # Bucle para estructurar e inyectar consecutivamente cada registro
                    for _, row_l in df_l_excel.iterrows():
                        if pd.notna(row_l['Nombre_Lider']) and str(row_l['Nombre_Lider']).strip() != "":
                            n_id_l = f"LID-{(len(st.session_state.BD_Lideres) + 1):02d}"
                            n_l_row = {
                                "ID_Lider": n_id_l, 
                                "Nombre_Lider": str(row_l['Nombre_Lider']).strip(), 
                                "Area": str(row_l['Area']).strip() if pd.notna(row_l['Area']) else "Metales", 
                                "Estatus": "Activo"
                            }
                            st.session_state.BD_Lideres = pd.concat([st.session_state.BD_Lideres, pd.DataFrame([n_l_row])], ignore_index=True)
                    
                    # Sincronizar y subir la base de datos de personal a GitHub
                    subir_excel_a_github("BD_Lideres.xlsx", st.session_state.BD_Lideres)
                    st.success("✅ ¡Personal integrado correctamente con folios LID y respaldado en GitHub!")
                    st.rerun()
            except Exception as e:
                st.error(f"Error al procesar el catálogo de personal: {e}")
                
        st.write("---")
        
        # 3. ALTA INDIVIDUAL TRADICIONAL Y EDICIÓN EN TABLA (CONSERVADA)
        with st.expander("➕ Dar de Alta Nuevo Líder Individual"):
            c_l1, c_l2 = st.columns(2)
            with c_l1: nuevo_nom = st.text_input("Nombre Completo:", key="txt_input_nuevo_lider_name")
            with c_l2: nueva_area = st.text_input("Área:", "Metales", key="txt_input_nuevo_lider_area")
            if st.button("➕ Registrar Líder Individual"):
                if nuevo_nom:
                    n_row = {"ID_Lider": f"LID-{(len(st.session_state.BD_Lideres) + 1):02d}", "Nombre_Lider": nuevo_nom, "Area": nueva_area, "Estatus": "Activo"}
                    st.session_state.BD_Lideres = pd.concat([st.session_state.BD_Lideres, pd.DataFrame([n_row])], ignore_index=True)
                    subir_excel_a_github("BD_Lideres.xlsx", st.session_state.BD_Lideres)
                    st.success("Líder registrado."); st.rerun()
                    
        st.write("📋 Catálogo Maestro de Líderes Autorizados (Editable):")
        df_l_edit = st.data_editor(st.session_state.BD_Lideres, use_container_width=True, hide_index=True, key="editor_catalogo_lideres_master")
        if st.button("💾 Sincronizar Cambios Manuales de Líderes"):
            st.session_state.BD_Lideres = df_l_edit
            subir_excel_a_github("BD_Lideres.xlsx", st.session_state.BD_Lideres)
            st.success("Catálogo de líderes sincronizado.")


    with tab3:
        st.subheader("🚨 Reset de Fábrica y Purga de Datos Controlada")
        metodo_purga = st.radio("Método de Purga:", ["❌ Purga Total Automática (Reset Completo)", "🔍 Seleccionar Registros Específicos para Eliminar"], horizontal=True, key="radio_metodo_purga_master")
        st.write("---")
        if metodo_purga == "❌ Purga Total Automática (Reset Completo)":
            st.error("⚠️ Peligro: Esta acción vaciará por completo todos los registros históricos en GitHub.")
            if st.checkbox("Confirmo que deseo aplicar un Reset de Fábrica Total.", key="chk_total_purga_final_u"):
                if st.button("🗑️ EJECUTAR PURGA MAESTRA TOTAL"):
                    st.session_state.BD_Tarimas = pd.DataFrame(columns=st.session_state.BD_Tarimas.columns)
                    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=st.session_state.BD_Detalle_Tarimas.columns)
                    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=st.session_state.BD_Datos_Generales_Remision.columns)
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success("💥 Sistema purgado por completo de forma masiva."); st.rerun()
        else:
            st.markdown("### 📦 1. Eliminar Tarimas del Inventario")
            if not st.session_state.BD_Tarimas.empty:
                df_tar_vista = st.session_state.BD_Tarimas.copy().drop(columns=["Es_Nueva"], errors="ignore")
                sel_tarimas = st.dataframe(df_tar_vista, use_container_width=True, on_select="rerun", selection_mode="multi-row", key="tabla_purga_tarimas_final")
                filas_tar = sel_tarimas.get("selection", {}).get("rows", [])
                if filas_tar:
                    ids_tar_eliminar = st.session_state.BD_Tarimas.iloc[filas_tar]['ID_Tarima'].tolist()
                    if st.button("🗑️ Eliminar Tarimas Seleccionadas"):
                        st.session_state.BD_Tarimas = st.session_state.BD_Tarimas[~st.session_state.BD_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                        st.session_state.BD_Detalle_Tarimas = st.session_state.BD_Detalle_Tarimas[~st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                        subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                        subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                        st.success("✅ Tarimas removidas."); st.rerun()
            else: st.write("No hay tarimas registradas en el inventario.")
            st.write("---")
            st.markdown("### 🚚 2. Eliminar Remisiones de Salida")
            if not st.session_state.BD_Datos_Generales_Remision.empty:
                sel_remisiones = st.dataframe(st.session_state.BD_Datos_Generales_Remision, use_container_width=True, on_select="rerun", selection_mode="multi-row", key="tabla_purga_remisiones_final")
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
                        st.success("✅ Remisiones seleccionadas purgadas con éxito."); st.rerun()
            else: st.write("No hay remisiones registradas en el historial.")
