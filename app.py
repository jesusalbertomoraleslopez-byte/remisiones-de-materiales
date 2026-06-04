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

# 1. CONFIGURACIÓN DE PÁGINA WEB RESPONSIVA
st.set_page_config(page_title="Remisiones de Materiales", layout="wide", page_icon="📦")

# Carga e inserción adaptativa del Banner Corporativo de 2000x400
try:
    banner_img = Image.open("REMISIONES APP.png")
    st.image(banner_img, use_container_width=True)
except FileNotFoundError:
    st.warning("⚠️ Cargando interfaz gráfica del banner superior...")
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
    st.session_state.BD_Tarimas = pd.DataFrame(columns=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus", "Es_Nueva"])

if "BD_Detalle_Tarimas" not in st.session_state:
    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Cantidad"])

if "BD_Datos_Generales_Remision" not in st.session_state:
    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=["ID_Remision", "Folio_Remision", "Fecha_Hora_Salida", "Nombre_Emisor", "Direccion_Emisor", "Nombre_Receptor", "Direccion_Receptor", "Tarimas_Asociadas"])
# 3. CONTROL DE SEGURIDAD MEDIANTE PASSWORDS MÚLTIPLES
st.sidebar.title("🔐 Control de Acceso")
admin_pass_input = st.sidebar.text_input("Contraseña Administrador:", type="password")
is_admin = (admin_pass_input == st.secrets.get("admin_password", ""))

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

# 4. CONECTOR DE PERSISTENCIA AUTOMÁTICA CON GITHUB API (REVISADO)
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
# 5. RENDERIZADO DE FORMATOS BAJO NORMA SIGRAMA FO-MET-10
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
    canvas.drawString(95, 37, "PROHIBIDA LA REPRODUCCIÓN TOTAL O PARCIAL SIN AUTORIZACIÓN DE INDUSTRIA SIGRAMA S.A. DE C.V.")
    canvas.restoreState()

def generar_pdf_tarima(id_tarima):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    tar_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == id_tarima].iloc[0]
    detalles = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == id_tarima]
    
    story.append(Spacer(1, 1.8 * inch))
    story.append(Paragraph(f"TARIMA<br/><br/><b>{id_tarima}</b>", ParagraphStyle('G', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1)))
    story.append(PageBreak())
    
    story.append(Paragraph(f"<b>Detalle Interno - Tarima {id_tarima}</b>", styles['Heading2']))
    story.append(Paragraph(f"<b>Operador:</b> {tar_info['Creado_Por']} | <b>Fecha:</b> {tar_info['Fecha_Creacion']}", styles['Normal']))
    story.append(Spacer(1, 0.3 * inch))
    for _, item in detalles.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
        story.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", styles['Normal']))
        story.append(Spacer(1, 0.4 * inch))
        story.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", ParagraphStyle('NG', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)))
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    return buffer.getvalue()

def generar_pdf_remision_general(datos_remision, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    style_b = ParagraphStyle('BB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_nt = ParagraphStyle('NT', parent=styles['Normal'], fontSize=9)
    
    t_header = Table([[Paragraph("LISTADO DE EMBARQUE", style_b)]], colWidths=[540])
    t_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#D32F2F")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header)
    
    panel = [
        [Paragraph("FOLIO:", styles['Normal']), Paragraph(str(datos_remision['Folio_Remision']), style_nt), Paragraph("FECHA:", styles['Normal']), Paragraph(str(datos_remision['Fecha_Hora_Salida']), style_nt)],
        [Paragraph("LÍDER:", styles['Normal']), Paragraph(str(datos_remision['Nombre_Emisor']), style_nt), Paragraph("ALMACÉN:", styles['Normal']), Paragraph(str(datos_remision['Direccion_Emisor']), style_nt)],
        [Paragraph("CLIENTE:", styles['Normal']), Paragraph(str(datos_remision['Nombre_Receptor']), style_nt), "", ""],
        [Paragraph("DESTINO:", styles['Normal']), Paragraph(str(datos_remision['Direccion_Receptor']), style_nt), "", ""]
    ]
    t_p = Table(panel, colWidths=[80, 190, 80, 190])
    t_p.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")), ('SPAN', (1,2), (3,2)), ('SPAN', (1,3), (3,3))]))
    story.append(t_p)
    story.append(Spacer(1, 0.2 * inch))
    
    mat_data = [[Paragraph("CANTIDAD", style_b), Paragraph("UNIDAD", style_b), Paragraph("CLAVE / MODELO", style_b), Paragraph("DESCRIPCIÓN", style_b), Paragraph("OBSERVACIONES", style_b)]]
    for _, r in df_detalles_remision.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == r['SKU']]
        nom = art.iloc[0]['Nombre'] if not art.empty else "Material"
        mat_data.append([Paragraph(str(int(r['Cantidad'])), style_nt), Paragraph("Piezas", style_nt), Paragraph(str(r['SKU']), style_nt), Paragraph(nom, style_nt), Paragraph(f"Tarima: {r['ID_Tarima']}", style_nt)])
    for _ in range(max(1, 6 - len(df_detalles_remision))): mat_data.append([Paragraph("", style_nt)] * 5)
    
    t_m = Table(mat_data, colWidths=[70, 60, 110, 160, 140])
    t_m.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575"))]))
    story.append(t_m)
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    return buffer.getvalue()

def generar_pdf_anexo_tarimas(lista_tarimas_id, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    style_b = ParagraphStyle('ABB', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=10)
    style_t = ParagraphStyle('ANT', parent=styles['Normal'], fontSize=9)
    
    for t_id in lista_tarimas_id:
        story.append(Spacer(1, 1.8 * inch))
        story.append(Paragraph(f"ANEXO: TARIMA<br/><b>{t_id}</b>", ParagraphStyle('C', parent=styles['Heading1'], fontSize=42, alignment=1)))
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
        t_d = Table(tabla_anexo, colWidths=[150, 240, 150])
        t_d.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F5F5F5")), ('GRID', (0,0), (-1,-1), 1, colors.white)]))
        story.append(t_d)
        story.append(PageBreak())
    if story: story.pop()
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    return buffer.getvalue()
# =============================================================================
# 6. MÓDULOS DE RENDERIZADO VISUAL EN PANTALLA
# =============================================================================
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
    else: st.info("No hay datos cargados para resumir por PO.")

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
    if not st.session_state.BD_Tarimas.empty and "Es_Nueva" not in st.session_state.BD_Tarimas.columns: st.session_state.BD_Tarimas["Es_Nueva"] = False

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
                st.success("¡Plantilla integrada y respaldada en GitHub!"); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
            
    if not st.session_state.BD_Tarimas.empty:
        st.write("---")
        df_estilado = st.session_state.BD_Tarimas.style.apply(lambda r: ['background-color: #FFF59D' if r['Es_Nueva'] else '' for _ in r], axis=1)
        sel_t = st.dataframe(df_estilado, use_container_width=True, column_order=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus"], on_select="rerun", selection_mode="multi-row")
        filas = sel_t.get("selection", {}).get("rows", [])
        if filas:
            elegidas = st.session_state.BD_Tarimas.iloc[filas]['ID_Tarima'].tolist()
            if len(elegidas) == 1:
                st.download_button(label=f"📥 Descargar PDF Tarima #{elegidas[0]}", data=generar_pdf_tarima(elegidas[0]), file_name=f"Tarima_{elegidas[0]}.pdf", mime="application/pdf")
            else:
                if st.button("📦 Unificar Lote de Impresión"):
                    buf_l = io.BytesIO()
                    doc_l = SimpleDocTemplate(buf_l, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
                    story_l, styles = [], getSampleStyleSheet()
                    for t_imp in elegidas:
                        det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == t_imp]
                        t_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == t_imp].iloc[0]
                        story_l.append(Spacer(1, 1.8 * inch)); story_l.append(Paragraph(f"TARIMA<br/><br/><b>{t_imp}</b>", ParagraphStyle('GL', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1))); story_l.append(PageBreak())
                        story_l.append(Paragraph(f"<b>Detalle Interno - Tarima {t_imp}</b>", styles['Heading2']))
                        story_l.append(Paragraph(f"<b>Operador:</b> {t_info['Creado_Por']} | <b>Fecha:</b> {t_info['Fecha_Creacion']}", styles['Normal']))
                        story_l.append(Spacer(1, 0.3 * inch))
                        for _, item in det.iterrows():
                            art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
                            nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
                            story_l.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", styles['Normal']))
                            story_l.append(Spacer(1, 0.4 * inch)); story_l.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", ParagraphStyle('NGL', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)))
                        story_l.append(PageBreak())
                    if story_l: story_l.pop()
                    doc_l.build(story_l, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
                    st.download_button(label="📥 Descargar Lote Completo (PDF)", data=buf_l.getvalue(), file_name="Lote_Tarimas.pdf", mime="application/pdf")

elif opcion_menu == "🚚 Módulo Remisiones":
    st.title("🚚 Generación de Remisiones de Salida")
    t_disp = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible']['ID_Tarima'].tolist()
    if not t_disp: st.warning("⚠️ No existen tarimas disponibles.")
    else:
        t_sel = st.multiselect("Seleccione Tarimas:", options=t_disp)
        col_e, col_r = st.columns(2)
        with col_e:
            nom_e = st.text_input("Líder / Emisor:", "Jesus Morales")
            dir_e = st.text_input("Almacén de Origen:", "Metales")
        with col_r:
            nom_r = st.text_input("Receptor / Cliente:", "Galvatec Industrias")
            dir_r = st.text_input("Dirección Destino:", "Prol. Valle Guadiana 919, Parque Industrial II, 35078 Gómez Palacio, Dgo.")
        if not is_admin: st.error("🔒 Operación Bloqueada.")
        else:
            if st.button("🚀 Confirmar Salida y Generar Nueva Remisión"):
                if not t_sel or not nom_r: st.error("Complete los campos.")
                else:
                    fol = f"E00{27 + len(st.session_state.BD_Datos_Generales_Remision)}"
                    reg = {"ID_Remision": len(st.session_state.BD_Datos_Generales_Remision)+1, "Folio_Remision": fol, "Fecha_Hora_Salida": datetime.datetime.now().strftime("%d/%m/%Y"), "Nombre_Emisor": nom_e, "Direccion_Emisor": dir_e, "Nombre_Receptor": nom_r, "Direccion_Receptor": dir_r, "Tarimas_Asociadas": t_sel}
                    st.session_state.BD_Datos_Generales_Remision = pd.concat([st.session_state.BD_Datos_Generales_Remision, pd.DataFrame([reg])], ignore_index=True)
                    st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(t_sel), 'Estatus'] = 'Remesada'
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success(f"✅ ¡Remisión {fol} Generada!"); st.rerun()
                    
        # Bloque de descargas de remisiones (se ejecuta si ya hay datos registrados)
    if not st.session_state.BD_Datos_Generales_Remision.empty:
        st.write("---")
        st.subheader("🖨️ Descarga Documental de Remisiones")
        r_sel = st.selectbox("Seleccione Folio para Descarga:", st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].unique())
        
        # Corrección de la línea cortada
        row = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'] == r_sel].iloc[0]
        df_det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(row['Tarimas_Asociadas'])]
        
        c1, c2 = st.columns(2)
        with c1: st.download_button(label="📥 Descargar Remisión (PDF)", data=generar_pdf_remision_general(row, df_det), file_name=f"Remision_{r_sel}.pdf", mime="application/pdf")
        with c2: st.download_button(label="📥 Descargar Anexo Tarimas (PDF)", data=generar_pdf_anexo_tarimas(row['Tarimas_Asociadas'], df_det), file_name=f"Anexo_{r_sel}.pdf", mime="application/pdf")
        # Manejo de descargas masivas en lote continuo
        if 'elegidas' in locals() and len(elegidas) > 1:
            if st.button("📦 Unificar y Preparar Lote de Impresión"):
                buf_l = io.BytesIO()
                doc_l = SimpleDocTemplate(buf_l, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
                story_l, styles = [], getSampleStyleSheet()
                for t_imp in elegidas:
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
                st.download_button(label="📥 Descargar Lote Completo (PDF)", data=buf_l.getvalue(), file_name="Lote_Tarimas_Sigrama.pdf", mime="application/pdf")
# =============================================================================
# MÓDULO FINAL: PANEL DE MANTENIMIENTO, EDICIÓN Y PURGA DE REGISTROS (SUPERUSER)
# =============================================================================
elif opcion_menu == "⚙️ Mantenimiento y Catálogos":
    st.title("⚙️ Panel de Mantenimiento Avanzado del Sistema")
    st.warning("⚠️ Acción Crítica: Las modificaciones realizadas impactan directamente en GitHub.")
    if "BD_Lideres" not in st.session_state:
        st.session_state.BD_Lideres = pd.DataFrame([{"ID_Lider": "LID-01", "Nombre_Lider": "Jesus Morales", "Area": "Metales", "Estatus": "Activo"}])

    tab1, tab2, tab3 = st.tabs(["📝 Ajustar Cantidades", "👤 Catálogo de Líderes", "🚨 Purga de Datos"])

    with tab1:
        st.subheader("✏️ Edición Rápida de Inventario (Detalle Tarimas)")
        if not st.session_state.BD_Detalle_Tarimas.empty:
            df_editable = st.data_editor(st.session_state.BD_Detalle_Tarimas, use_container_width=True, disabled=["ID_Detalle", "ID_Tarima", "SKU", "PO"], hide_index=True)
            if st.button("💾 Guardar Cambios de Cantidades en GitHub"):
                st.session_state.BD_Detalle_Tarimas = df_editable
                subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                st.success("✅ Cantidades corregidas y sincronizadas con éxito."); st.rerun()
        else: st.info("No hay registros en el detalle de tarimas para modificar.")

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

    # --- PESTAÑA 3: PURGA SELECCIONADA Y MASIVA DE DATOS (ACTUALIZADO CON CASILLAS) ---
    with tab3:
        st.subheader("🚨 Reset de Fábrica y Purga de Datos Controlada")
        st.write("Seleccione el método de purga que desea aplicar sobre las bases de datos maestras:")
        
        metodo_purga = st.radio(
            "Método de Purga:",
            ["❌ Purga Total Automática (Reset Completo)", "🔍 Seleccionar Registros Específicos para Eliminar"],
            horizontal=True
        )
        
        st.write("---")
        
        # OPCIÓN 1: VACIADO ABSOLUTO AUTOMÁTICO
        if metodo_purga == "❌ Purga Total Automática (Reset Completo)":
            st.error("⚠️ Peligro: Esta acción vaciará por completo todos los registros históricos del sistema y limpiará los archivos de GitHub a ceros.")
            confirmar_total = st.checkbox("Confirmo que deseo aplicar un Reset de Fábrica Total.", key="chk_total")
            
            if confirmar_total:
                if st.button("🗑️ EJECUTAR PURGA MAESTRA TOTAL"):
                    st.session_state.BD_Tarimas = pd.DataFrame(columns=st.session_state.BD_Tarimas.columns)
                    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=st.session_state.BD_Detalle_Tarimas.columns)
                    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=st.session_state.BD_Datos_Generales_Remision.columns)
                    
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success("💥 Sistema y repositorio purgados por completo a ceros de forma masiva."); st.rerun()

        # OPCIÓN 2: SELECCIÓN DETALLADA MEDIANTE CASILLAS DE VERIFICACIÓN NATIVAS
        else:
            st.info("💡 Consejo: Seleccione una o más casillas en las tablas inferiores para habilitar los botones de remoción parcial.")
            
            # --- SUBSECCIÓN A: SELECCIÓN DE TARIMAS ---
            st.markdown("### 📦 1. Eliminar Tarimas del Inventario")
            if not st.session_state.BD_Tarimas.empty:
                # Quitamos la columna técnica de color para no ensuciar la tabla visual
                df_tar_vista = st.session_state.BD_Tarimas.copy().drop(columns=["Es_Nueva"], errors="ignore")
                
                sel_tarimas = st.dataframe(
                    df_tar_vista,
                    use_container_width=True,
                    on_select="rerun",
                    selection_mode="multi-row",
                    key="tabla_purga_tarimas"
                )
                filas_tar = sel_tarimas.get("selection", {}).get("rows", [])
                
                if filas_tar:
                    ids_tar_eliminar = st.session_state.BD_Tarimas.iloc[filas_tar]['ID_Tarima'].tolist()
                    st.warning(f"⚠️ Seleccionadas para eliminar: {', '.join(ids_tar_eliminar)}")
                    
                    if st.button("🗑️ Eliminar Tarimas Seleccionadas"):
                        # Filtrar y remover en caliente conservando el resto de registros
                        st.session_state.BD_Tarimas = st.session_state.BD_Tarimas[~st.session_state.BD_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                        # Purga en cascada opcional para limpiar sus respectivos detalles relacionales
                        st.session_state.BD_Detalle_Tarimas = st.session_state.BD_Detalle_Tarimas[~st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                        
                        # Actualizar repositorios en GitHub
                        subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                        subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                        st.success("✅ Registros de Tarimas removidos e inyectados en GitHub."); st.rerun()
            else:
                st.write("No hay tarimas registradas en el inventario activo.")
                
            st.write("---")
            
            # --- SUBSECCIÓN B: SELECCIÓN DE REMISIONES ---
            st.markdown("### 🚚 2. Eliminar Remisiones de Salida")
            if not st.session_state.BD_Datos_Generales_Remision.empty:
                sel_remisiones = st.dataframe(
                    st.session_state.BD_Datos_Generales_Remision,
                    use_container_width=True,
                    on_select="rerun",
                    selection_mode="multi-row",
                    key="tabla_purga_remisiones"
                )
                filas_rem = sel_remisiones.get("selection", {}).get("rows", [])
                
                if filas_rem:
                    ids_rem_eliminar = st.session_state.BD_Datos_Generales_Remision.iloc[filas_rem]['Folio_Remision'].tolist()
                    # Rescatamos los IDs de las tarimas de esas remisiones para devolverles su estatus 'Disponible'
                    tarimas_afectadas = []
                    for idx in filas_rem:
                        tarimas_afectadas.extend(st.session_state.BD_Datos_Generales_Remision.iloc[idx]['Tarimas_Asociadas'])
                        
                    st.warning(f"⚠️ Seleccionadas para eliminar: {', '.join(ids_rem_eliminar)}")
                    
                    if st.button("🗑️ Eliminar Remisiones Seleccionadas"):
                        # Restaurar el estatus de las tarimas asociadas a 'Disponible' antes de borrar la remisión
                        st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(tarimas_afectadas), 'Estatus'] = 'Disponible'
                        
                        # Remover físicamente el folio de la remisión
                        st.session_state.BD_Datos_Generales_Remision = st.session_state.BD_Datos_Generales_Remision[~st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].isin(ids_rem_eliminar)]
                        
                        # Sincronizar bases de datos afectadas con GitHub
                        subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                        subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                        st.success("✅ Folios de remisión eliminados y tarimas restauradas a Disponible."); st.rerun()
            else:
                st.write("No hay folios de remisiones emitidos en el historial.")

