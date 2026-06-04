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

# DESPLIEGUE DEL BANNER CORPORATIVO LOCAL
try:
    banner_img = Image.open("REMISIONES APP.png")
    st.image(banner_img, use_container_width=True)
except FileNotFoundError:
    st.warning("⚠️ Cargando interfaz gráfica del banner...")
st.write("")
# 2. MODELO RELACIONAL EN MEMORIA (PANDAS + SESSION STATE)
if "BD_Articulos" not in st.session_state:
    st.session_state.BD_Articulos = pd.DataFrame([
        {"SKU": "12-B-9016-01", "Nombre": "Lámina Galvanizada Sigrama", "Calibre_Espesor": "Calibre 22", "Dimensiones_Pieza": "3x10 ft", "Acabado_Superficial": "Zintro"},
        {"SKU": "SKU-002", "Nombre": "Placa de Acero Comercial", "Calibre_Espesor": "1/4 pulgada", "Dimensiones_Pieza": "4x8 ft", "Acabado_Superficial": "Negro"}
    ])

if "BD_Ordenes_Compra" not in st.session_state:
    st.session_state.BD_Ordenes_Compra = pd.DataFrame([
        {"PO": "PO-10001", "SKU": "12-B-9016-01", "Cantidad_Solicitada": 500}
    ])
if "BD_Tarimas" not in st.session_state:
    st.session_state.BD_Tarimas = pd.DataFrame(columns=[
        "ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus", "Es_Nueva"
    ])

if "BD_Detalle_Tarimas" not in st.session_state:
    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Cantidad"])

if "BD_Datos_Generales_Remision" not in st.session_state:
    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=[
        "ID_Remision", "Folio_Remision", "Fecha_Hora_Salida", "Nombre_Emisor", "Direccion_Emisor", "Nombre_Receptor", "Direccion_Receptor", "Tarimas_Asociadas"
    ])
# 3. SISTEMA DE SEGURIDAD (SECRETS) Y NAVEGACIÓN
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

st.sidebar.title("🧭 Navegación")
opcion_menu = st.sidebar.radio(
    "Seleccione un Módulo:",
    ["📊 Dashboard e Históricos", "🔍 Centro de Consultas", "📦 Módulo Tarimas", "🚚 Módulo Remisiones"]
)
# 4. CONECTOR DE PERSISTENCIA AUTOMÁTICA CON GITHUB API
def subir_excel_a_github(file_name, dataframe_to_save):
    try:
        GITHUB_TOKEN = st.secrets["github_token"]
        REPO_OWNER = "jesusalbertomoraleslopez-byte"
        REPO_NAME = "remisiones-de-materiales"
        BRANCH = "main"
        
        buffer_git = io.BytesIO()
        with pd.ExcelWriter(buffer_git, engine='openpyxl') as writer:
            dataframe_to_save.to_excel(writer, index=False, sheet_name='Datos_Sistema')
        
        base64_content = base64.b64encode(buffer_git.getvalue()).decode("utf-8")
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{file_name}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
        
        res_get = requests.get(url, headers=headers)
        sha = res_get.json().get("sha") if res_get.status_code == 200 else None
            
        payload = {"message": f"Sincronización: {file_name}", "content": base64_content, "branch": BRANCH}
        if sha: payload["sha"] = sha
            
        res_put = requests.put(url, json=payload, headers=headers)
        return res_put.status_code in [200, 201]
    except Exception:
        return False
# 5. MOTOR DE RENDERIZADO PDF (FORMATO CORPORATIVO SIGRAMA)
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
    style_n, style_ng = styles['Normal'], ParagraphStyle('NG', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)
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
        [Paragraph("LÍDER:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Emisor']), style_normal_text), Paragraph("ALMACÉN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Emisor']), style_normal_text)],
        [Paragraph("DESTINO:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Receptor']), style_normal_text), "", ""],
        [Paragraph("DIRECCIÓN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Receptor']), style_normal_text), "", ""]
    ]
    t_panel = Table(datos_panel, colWidths=[100, 170, 90, 180])
    t_panel.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")), ('SPAN', (1,2), (3,2)), ('SPAN', (1,3), (3,3))]))
    story.append(t_panel)
    story.append(Spacer(1, 0.2 * inch))
    tabla_materiales = [[Paragraph("CANTIDAD", style_blanco_bold), Paragraph("UNIDAD", style_blanco_bold), Paragraph("CLAVE", style_blanco_bold), Paragraph("DESCRIPCIÓN", style_blanco_bold), Paragraph("OBSERVACIONES", style_blanco_bold)]]
    for _, row in df_detalles_remision.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == row['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Material"
        tabla_materiales.append([Paragraph(str(int(row['Cantidad'])), style_normal_text), Paragraph("Piezas", style_normal_text), Paragraph(str(row['SKU']), style_normal_text), Paragraph(nom_art, style_normal_text), Paragraph(f"Tarima: #{row['ID_Tarima']}", style_normal_text)])
    for _ in range(max(1, 6 - len(df_detalles_remision))): tabla_materiales.append([Paragraph("", style_normal_text)] * 5)
    t_mat = Table(tabla_materiales, colWidths=[70, 60, 100, 180, 130])
    t_mat.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575"))]))
    story.append(t_mat)
    story.append(Spacer(1, 0.3 * inch))
    t_firmas = Table([[Paragraph("ENTREGA", style_normal_bold), Paragraph("RECIBE", style_normal_bold)]], colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('LINEBELOW', (0,0), (0,0), 1, colors.black), ('LINEBELOW', (1,0), (1,0), 1, colors.black), ('TOPPADDING', (0,0), (-1,-1), 25)]))
    story.append(t_firmas)
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer.getvalue()
def generar_pdf_anexo_tarimas(lista_tarimas_id, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    style_b = ParagraphStyle('ABB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_t = ParagraphStyle('ANT', parent=styles['Normal'], fontSize=9)
    for t_id in lista_tarimas_id:
        style_c = ParagraphStyle('C', parent=styles['Heading1'], fontSize=42, alignment=1)
        story.append(Spacer(1, 1.8 * inch))
        story.append(Paragraph(f"ANEXO: TARIMA #{t_id}", style_c))
        story.append(Spacer(1, 0.4 * inch))
        t_bar = Table([["|||||||||||||||||||||||||||||||"], [f"*TARIMA-{t_id}*"]], colWidths=[540])
        t_bar.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('TEXTCOLOR', (0,0), (-1,1), colors.darkgray)]))
        story.append(t_bar)
        story.append(PageBreak())
        story.append(Paragraph(f"<b>DETALLE ESPECÍFICO - TARIMA {t_id}</b>", styles['Heading2']))
        sub_det = df_detalles_remision[df_detalles_remision['ID_Tarima'] == t_id]
        tabla_anexo = [[Paragraph("PO ASOCIADA", style_b), Paragraph("SKU / PRODUCTO", style_b), Paragraph("CANTIDAD", style_b)]]
        for _, s_row in sub_det.iterrows():
            tabla_anexo.append([Paragraph(str(s_row['PO']), style_t), Paragraph(str(s_row['SKU']), style_t), Paragraph(str(int(s_row['Cantidad'])), style_t)])
        t_det = Table(tabla_anexo, colWidths=[150, 240, 150])
        t_det.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F5F5F5")), ('GRID', (0,0), (-1,-1), 1, colors.white)]))
        story.append(t_det)
        story.append(PageBreak())
    if story: story.pop()
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer.getvalue()
# 6. RENDERIZADO DE LAS VISTAS ACTIVAS
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
    m1.metric("📦 Total Tarimas", f"{t_tar} Disp.")
    m2.metric("🟢 Tarimas Disponibles", disp)
    m3.metric("🚚 Tarimas Remesadas", rem)
    m4.metric("🔢 Total Piezas en Almacén", f"{total_piezas:,} PZS")
    
    st.write("---")
    st.subheader("📋 Resumen de Cumplimiento por Pedido (PO)")
    if not st.session_state.BD_Detalle_Tarimas.empty:
        df_c = pd.merge(st.session_state.BD_Detalle_Tarimas, st.session_state.BD_Tarimas[['ID_Tarima', 'Estatus']], on='ID_Tarima', how='left')
        resumen_po = df_c.groupby('PO').agg(
            Cant_Tarimas=('ID_Tarima', 'nunique'), Total_Piezas=('Cantidad', 'sum'),
            Piezas_Disp=('Cantidad', lambda x: x[df_c.loc[x.index, 'Estatus'] == 'Disponible'].sum()),
            Piezas_Rem=('Cantidad', lambda x: x[df_c.loc[x.index, 'Estatus'] == 'Remesada'].sum())
        ).reset_index()
        resumen_po.columns = ["Orden de Compra (PO)", "Cant. Tarimas", "Total Piezas", "Piezas Disponibles", "Piezas Remesadas"]
        st.dataframe(resumen_po, use_container_width=True, hide_index=True)
    else:
        st.info("No hay datos cargados para resumir por PO.")
elif opcion_menu == "🔍 Centro de Consultas":
    st.title("🔍 Centro de Consultas Avanzado")
    col_sel, col_val = st.columns(2)
    with col_sel: tipo_filtro = st.selectbox("Filtrar por:", ["Ninguno", "SKU", "PO", "Folio_Remision"])
    with col_val: valor_filtro = st.text_input("Término de búsqueda:")
    res = st.session_state.BD_Detalle_Tarimas.copy()
    if tipo_filtro == "SKU" and valor_filtro: res = res[res['SKU'].str.contains(valor_filtro, case=False)]
    elif tipo_filtro == "PO" and valor_filtro: res = res[res['PO'].str.contains(valor_filtro, case=False)]
    elif tipo_filtro == "Folio_Remision" and valor_filtro:
        rem_m = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].str.contains(valor_filtro, case=False)]
        tar_v = [t for r in rem_m['Tarimas_Asociadas'] for t in r]
        res = res[res['ID_Tarima'].isin(tar_v)]
    st.dataframe(res, use_container_width=True)
elif opcion_menu == "📦 Módulo Tarimas":
    st.title("📦 Carga de Tarimas")
    df_p = pd.DataFrame([{"Tarima": "Bulto_1", "Producto/SKU": "12-B-9016-01", "PO": "PO-10001", "Cantidad": 191}])
    buf_p = io.BytesIO()
    with pd.ExcelWriter(buf_p, engine='openpyxl') as wr: df_p.to_excel(wr, index=False)
    st.download_button(label="📥 Descargar Plantilla de Ejemplo (.xlsx)", data=buf_p.getvalue(), file_name="plantilla_tarimas.xlsx")
    
    if not st.session_state.BD_Tarimas.empty and "Es_Nueva" not in st.session_state.BD_Tarimas.columns:
        st.session_state.BD_Tarimas["Es_Nueva"] = False

    if not is_admin: st.error("🔒 Área Bloqueada: Requiere contraseña de Administrador.")
    else:
        st.success("🔓 Acceso Autorizado.")
        arch = st.file_uploader("Suba la Plantilla Excel", type=["xlsx"])
        col_t1, col_t2 = st.columns(2)
        with col_t1: tipo_t = st.selectbox("Tipo:", ["Cuadrada", "Rectangular"])
        with col_t2: oper = st.text_input("Líder:", "Jesus Morales")
        if arch and st.button("Procesar e Integrar Plantilla"):
            try:
                df_ex = pd.read_excel(arch)
                if not st.session_state.BD_Tarimas.empty: st.session_state.BD_Tarimas["Es_Nueva"] = False
                for t_orig in df_ex['Tarima'].unique():
                    n_id_tpm = f"TPM-{(len(st.session_state.BD_Tarimas) + 1):04d}"
                    n_t = {"ID_Tarima": n_id_tpm, "Tarima_Origen_Excel": t_orig, "Fecha_Creacion": datetime.datetime.now().strftime("%d/%m/%Y"), "Ubicacion_Actual": "Metales", "Creado_Por": oper, "Tipo_Tarima": tipo_t, "Estatus": "Disponible", "Es_Nueva": True}
                    st.session_state.BD_Tarimas = pd.concat([st.session_state.BD_Tarimas, pd.DataFrame([n_t])], ignore_index=True)
                    items = df_ex[df_ex['Tarima'] == t_orig]
                    for _, item in items.iterrows():
                        st.session_state.BD_Detalle_Tarimas = pd.concat([st.session_state.BD_Detalle_Tarimas, pd.DataFrame([{"ID_Detalle": len(st.session_state.BD_Detalle_Tarimas)+1, "ID_Tarima": n_id_tpm, "SKU": item['Producto/SKU'], "PO": item['PO'], "Cantidad": item['Cantidad']}])], ignore_index=True)
                subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                st.success("¡Plantilla integrada y respaldada en GitHub!")
            except Exception as e: st.error(f"Error: {e}")
            
    if not st.session_state.BD_Tarimas.empty:
        st.write("---")
        df_estilado = st.session_state.BD_Tarimas.style.apply(lambda r: ['background-color: #FFF59D' if r['Es_Nueva'] else '' for _ in r], axis=1)
        sel_t = st.dataframe(df_estilado, use_container_width=True, column_order=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus"], on_select="rerun", selection_mode="multi-row")
        filas = sel_t.get("selection", {}).get("rows", [])
        if filas:
            elegidas = st.session_state.BD_Tarimas.iloc[filas]['ID_Tarima'].tolist()
            if len(elegidas) == 1:
                st.download_button(label=f"📥 Descargar PDF Tarima #{elegidas[0]}", data=generar_pdf_tarima(elegidas[0]), file_name=f"Tarima_{elegidas[0]}.pdf")
            else:
                if st.button("📦 Unificar y Preparar Lote de Impresión"):
                    buf_l = io.BytesIO()
                    doc_l = SimpleDocTemplate(buf_l, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
                    story_l, styles = [], getSampleStyleSheet()
                    for t_imp in elegidas:
                        det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == t_imp]
                        t_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == t_imp].iloc[0]
                        story_l.append(Spacer(1, 1.8 * inch)); story_l.append(Paragraph(f"TARIMA<br/><br/><b>#{t_imp}</b>", ParagraphStyle('G_L', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1))); story_l.append(PageBreak())
                        story_l.append(Paragraph(f"<b>Detalle Interno - Tarima #{t_imp}</b>", styles['Heading2']))
                        story_l.append(Paragraph(f"<b>Operador:</b> {t_info['Creado_Por']} | <b>Fecha:</b> {t_info['Fecha_Creacion']}", styles['Normal']))
                        story_l.append(Spacer(1, 0.3 * inch))
                        for _, item in det.iterrows():
                            art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
                            nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
                            story_l.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", styles['Normal']))
                            story_l.append(Spacer(1, 0.4 * inch)); story_l.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", ParagraphStyle('NG_L', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)))
                        story_l.append(PageBreak())
                    if story_l: story_l.pop()
                    doc_l.build(story_l, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
                    st.download_button(label="📥 Descargar Lote Completo (PDF)", data=buf_l.getvalue(), file_name="Lote_Tarimas.pdf")
        else: st.warning("Seleccione filas en la tabla para descargar.")
        # Entrada de datos logísticos para el encabezado del documento formal
        col_e, col_r = st.columns(2)
        with col_e:
            nom_e = st.text_input("Líder / Emisor:", "Jesus Morales")
            dir_e = st.text_input("Almacén de Origen:", "Metales")
        with col_r:
            nom_r = st.text_input("Receptor / Cliente:", "Galvatec Industrias")
            dir_r = st.text_input(
                "Dirección Destino:", 
                "Prol. Valle Guadiana 919, Parque Industrial II, 35078 Gómez Palacio, Dgo."
            )
        # Bloque de validación y acción restringida para el Administrador
        if not is_admin: 
            st.error("🔒 Operación Bloqueada: Ingrese la contraseña de Administrador para confirmar.")
        else:
            if st.button("🚀 Confirmar Salida y Generar Nueva Remisión"):
                if not t_sel or not nom_r: 
                    st.error("Por favor, complete todos los campos obligatorios antes de continuar.")
                else:
                    # Lógica de cálculo de folios automáticos del sistema
                    fol = f"E00{27 + len(st.session_state.BD_Datos_Generales_Remision)}"
                    reg = {
                        "ID_Remision": len(st.session_state.BD_Datos_Generales_Remision) + 1, 
                        "Folio_Remision": fol, 
                        "Fecha_Hora_Salida": datetime.datetime.now().strftime("%d/%m/%Y"), 
                        "Nombre_Emisor": nom_e, "Direccion_Emisor": dir_e, 
                        "Nombre_Receptor": nom_r, "Direccion_Receptor": dir_r, 
                        "Tarimas_Asociadas": t_sel
                    }
                    
                    # Actualización de estados relacionales en la memoria caché
                    st.session_state.BD_Datos_Generales_Remision = pd.concat(
                        [st.session_state.BD_Datos_Generales_Remision, pd.DataFrame([reg])], 
                        ignore_index=True
                    )
                    st.session_state.BD_Tarimas.loc[
                        st.session_state.BD_Tarimas['ID_Tarima'].isin(t_sel), 'Estatus'
                    ] = 'Remesada'
                    
                    # PERSISTENCIA EN GITHUB: Inyección física e inmediata de los archivos actualizados
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success(f"✅ ¡Remisión {fol} Generada y Respaldada Exitosamente en GitHub!")
    # Módulo de descargas históricas (Visible si ya existen registros en el sistema)
    if not st.session_state.BD_Datos_Generales_Remision.empty:
        st.write("---")
        st.subheader("🖨️ Descarga Documental de Remisiones")
        
        # Selector dinámico de Folio emitido
        r_sel = st.selectbox(
            "Seleccione Folio para Descarga:", 
            st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].unique()
        )
        
        # Recuperación de registros y cruce de datos
        row = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'] == r_sel].iloc[0]
        df_det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(row['Tarimas_Asociadas'])]
        
        # Botones asíncronos finales para la obtención del PDF
        c1, c2 = st.columns(2)
        with c1: 
            st.download_button(
                label="📥 Descargar Remisión (PDF)", 
                data=generar_pdf_remision_general(row, df_det), 
                file_name=f"Remision_{r_sel}.pdf", 
                mime="application/pdf"
            )
        with c2: 
            st.download_button(
                label="📥 Descargar Anexo Tarimas (PDF)", 
                data=generar_pdf_anexo_tarimas(row['Tarimas_Asociadas'], df_det), 
                file_name=f"Anexo_{r_sel}.pdf", 
                mime="application/pdf"
            )
