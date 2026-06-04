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

# 1. CONFIGURACIÓN DE LA PÁGINA WEB
st.set_page_config(page_title="Remisiones de Materiales", layout="wide", page_icon="📦")

# BANNER SUPERIOR CORPORATIVO RESPONSIBO
try:
    banner_img = Image.open("REMISIONES APP.png")
    st.image(banner_img, use_container_width=True)
except FileNotFoundError:
    st.warning("⚠️ Cargando interfaz gráfica del banner...")
st.write("")
# =============================================================================
# CONFIGURACIÓN ESTRUCTURAL DE ARQUITECTURA DE DATOS (GITHUB)
# =============================================================================

# CONFIGURA AQUÍ LOS DATOS DE TU REPOSITORIO DE FORMA DIRECTA:
REPO_OWNER = "jesusalbertomoraleslopez-byte"   # Tu usuario de GitHub
REPO_NAME = "remisiones-de-materiales"          # El nombre de tu repositorio
BRANCH = "main"                                  # Tu rama principal

# El Token de seguridad NO se escribe aquí por seguridad. 
# Se lee automáticamente desde tus "Secrets" de Streamlit Cloud usando la siguiente llave:
def obtener_token_git():
    try:
        return st.secrets["github_token"]
    except KeyError:
        return None

def descargar_excel_de_github(file_name):
    """Descarga un archivo Excel desde el repositorio de GitHub y lo convierte a DataFrame."""
    token = obtener_token_git()
    if not token:
        return None
    url = f"https://github.com{REPO_OWNER}/{REPO_NAME}/contents/{file_name}?ref={BRANCH}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        try:
            content_b64 = response.json().get("content", "")
            content_bytes = base64.b64decode(content_b64)
            return pd.read_excel(io.BytesIO(content_bytes))
        except Exception:
            return None
    return None

def subir_excel_a_github(file_name, df_to_save):
    """Convierte un DataFrame a Excel y lo guarda permanentemente en tu repositorio de GitHub."""
    token = obtener_token_git()
    if not token:
        st.sidebar.error("❌ Token de GitHub no configurado en Secrets.")
        return False
        
    url = f"https://github.com{REPO_OWNER}/{REPO_NAME}/contents/{file_name}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
    
    # Convertir DataFrame a binario Excel
    buffer_git = io.BytesIO()
    with pd.ExcelWriter(buffer_git, engine='openpyxl') as writer:
        df_to_save.to_excel(writer, index=False, sheet_name='Datos_Sistema')
    content_bytes = buffer_git.getvalue()
    base64_content = base64.b64encode(content_bytes).decode("utf-8")
    
    # Obtener SHA si el archivo ya existe para poder sobreescribirlo
    res_get = requests.get(f"{url}?ref={BRANCH}", headers=headers)
    sha = res_get.json().get("sha") if res_get.status_code == 200 else None
    
    payload = {
        "message": f"Actualización automática desde App Streamlit: {file_name}",
        "content": base64_content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha
        
    res_put = requests.put(url, json=payload, headers=headers)
    if res_put.status_code in [200, 201]:
        st.sidebar.success(f"💾 Guardado permanente en GitHub: {file_name}")
        return True
    return False
# Inicializar Catálogo de Artículos Base (Estático en código)
if "BD_Articulos" not in st.session_state:
    st.session_state.BD_Articulos = pd.DataFrame([
        {"SKU": "12-B-9016-01", "Nombre": "Lámina Galvanizada Sigrama", "Calibre_Espesor": "Calibre 22", "Dimensiones_Pieza": "3x10 ft", "Acabado_Superficial": "Zintro"},
        {"SKU": "SKU-002", "Nombre": "Placa de Acero Comercial", "Calibre_Espesor": "1/4 pulgada", "Dimensiones_Pieza": "4x8 ft", "Acabado_Superficial": "Negro"}
    ])

# CARGA DINÁMICA DESDE GITHUB PARA EVITAR PÉRDIDA DE DATOS AL REINICIAR
if "BD_Tarimas" not in st.session_state:
    df_cloud = descargar_excel_de_github("BD_Tarimas.xlsx")
    st.session_state.BD_Tarimas = df_cloud if df_cloud is not None else pd.DataFrame(columns=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus", "Es_Nueva"])

if "BD_Detalle_Tarimas" not in st.session_state:
    df_cloud = descargar_excel_de_github("BD_Detalle_Tarimas.xlsx")
    st.session_state.BD_Detalle_Tarimas = df_cloud if df_cloud is not None else pd.DataFrame(columns=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Cantidad"])

if "BD_Datos_Generales_Remision" not in st.session_state:
    df_cloud = descargar_excel_de_github("BD_Datos_Generales_Remision.xlsx")
    if df_cloud is not None:
        # Re-convertir la cadena de texto de la lista a tipo lista real de Python
        df_cloud["Tarimas_Asociadas"] = df_cloud["Tarimas_Asociadas"].apply(lambda x: eval(str(x)) if isinstance(x, str) else x)
        st.session_state.BD_Datos_Generales_Remision = df_cloud
    else:
        st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=["ID_Remision", "Folio_Remision", "Fecha_Hora_Salida", "Nombre_Emisor", "Direccion_Emisor", "Nombre_Receptor", "Direccion_Receptor", "Tarimas_Asociadas"])
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
    style_g = ParagraphStyle('G', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1)
    story.append(Spacer(1, 1.8 * inch)); story.append(Paragraph(f"TARIMA<br/><br/><b>#{id_tarima}</b>", style_g)); story.append(PageBreak())
    style_n, style_ng = styles['Normal'], ParagraphStyle('NG', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)
    story.append(Paragraph(f"<b>Detalle Interno - Tarima #{id_tarima}</b>", styles['Heading2']))
    story.append(Paragraph(f"<b>Operador:</b> {tarima_info['Creado_Por']} | <b>Fecha:</b> {tarima_info['Fecha_Creacion']}", style_n))
    story.append(Spacer(1, 0.3 * inch))
    for _, item in detalles.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
        story.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", style_n))
        story.append(Spacer(1, 0.4 * inch)); story.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", style_ng))
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer

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
        tabla_materiales.append([Paragraph(str(int(row['Cantidad'])), style_normal_text), Paragraph("Piezas", style_normal_text), Paragraph(str(row['SKU']), style_normal_text), Paragraph(nom_art, style_normal_text), Paragraph(f"Tarima: #{row['ID_Tarima']}", style_normal_text)])
    for _ in range(max(1, 8 - len(df_detalles_remision))):
        tabla_materiales.append([Paragraph("", style_normal_text)] * 5)
    t_mat = Table(tabla_materiales, colWidths=[1.0 * inch, 0.8 * inch, 1.5 * inch, 2.7 * inch, 1.5 * inch])
    t_mat.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575"))]))
    story.append(t_mat)
    story.append(Spacer(1, 0.3 * inch))
    t_firmas = Table([[Paragraph("ENTREGA", style_normal_bold), Paragraph("RECIBE", style_normal_bold)]], colWidths=[3.75 * inch, 3.75 * inch])
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
    for t_id in lista_tarimas_id:
        style_c = ParagraphStyle('C', parent=styles['Heading1'], fontSize=42, alignment=1)
        story.append(Spacer(1, 1.8 * inch)); story.append(Paragraph(f"ANEXO: TARIMA #{t_id}", style_c)); story.append(Spacer(1, 0.4 * inch))
        t_bar = Table([["|||||||||||||||||||||||||||||||"], [f"*TARIMA-{t_id}*"]], colWidths=[7.5 * inch])
        t_bar.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('TEXTCOLOR', (0,0), (-1,1), colors.darkgray)]))
        story.append(t_bar); story.append(PageBreak())
        story.append(Paragraph(f"<b>DETALLE ESPECÍFICO - TARIMA {t_id}</b>", styles['Heading2']))
        sub_det = df_detalles_remision[df_detalles_remision['ID_Tarima'] == t_id]
        tabla_anexo = [[Paragraph("PO ASOCIADA", style_b), Paragraph("SKU / PRODUCTO", style_b), Paragraph("CANTIDAD", style_b)]]
        for _, s_row in sub_det.iterrows():
            tabla_anexo.append([Paragraph(str(s_row['PO']), style_t), Paragraph(str(s_row['SKU']), style_t), Paragraph(str(int(s_row['Cantidad'])), style_t)])
        t_det = Table(tabla_anexo, colWidths=[2.5 * inch, 2.5 * inch, 2.5 * inch])
        t_det.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F5F5F5")), ('GRID', (0,0), (-1,-1), 1, colors.white)]))
        story.append(t_det); story.append(PageBreak())
    if story: story.pop()
    doc.build(story, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
    buffer.seek(0)
    return buffer
# =============================================================================
# SECCIÓN 7A: MÓDULO TARIMAS - PROCESAMIENTO E INTEGRACIÓN CON GITHUB
# =============================================================================
elif opcion_menu == "📦 Módulo Tarimas":
    st.title("📦 Carga de Tarimas")
    st.subheader("📋 Formato Requerido")
    
    # Descarga de la plantilla muestra en memoria
    df_p = pd.DataFrame([{"Tarima": "Bulto_1", "Producto/SKU": "12-B-9016-01", "PO": "PO-10001", "Cantidad": 191}])
    buf_p = io.BytesIO()
    with pd.ExcelWriter(buf_p, engine='openpyxl') as wr: 
        df_p.to_excel(wr, index=False)
    st.download_button(label="📥 Descargar Plantilla de Ejemplo (.xlsx)", data=buf_p.getvalue(), file_name="plantilla_tarimas.xlsx")
    
    if not st.session_state.BD_Tarimas.empty and "Es_Nueva" not in st.session_state.BD_Tarimas.columns:
        st.session_state.BD_Tarimas["Es_Nueva"] = False

    if not is_admin: 
        st.error("🔒 Área Bloqueada: Requiere contraseña de Administrador para subir datos.")
    else:
        st.success("🔓 Acceso Autorizado para Operaciones.")
        arch = st.file_uploader("Suba la Plantilla Excel Rellenada", type=["xlsx"])
        col_t1, col_t2 = st.columns(2)
        with col_t1: tipo_t = st.selectbox("Tipo de Tarima:", ["Cuadrada", "Rectangular"])
        with col_t2: oper = st.text_input("Líder Operativo:", "Jesus Morales")
        
        if arch and st.button("Procesar e Integrar Plantilla"):
            try:
                df_ex = pd.read_excel(arch)
                if not st.session_state.BD_Tarimas.empty: 
                    st.session_state.BD_Tarimas["Es_Nueva"] = False
                    
                for t_orig in df_ex['Tarima'].unique():
                    num_consecutivo = len(st.session_state.BD_Tarimas) + 1
                    nuevo_id_tpm = f"TPM-{num_consecutivo:04d}"
                    
                    n_t = {"ID_Tarima": nuevo_id_tpm, "Tarima_Origen_Excel": t_orig, "Fecha_Creacion": datetime.datetime.now().strftime("%d/%m/%Y"), "Ubicacion_Actual": "Metales", "Creado_Por": oper, "Tipo_Tarima": tipo_t, "Estatus": "Disponible", "Es_Nueva": True}
                    st.session_state.BD_Tarimas = pd.concat([st.session_state.BD_Tarimas, pd.DataFrame([n_t])], ignore_index=True)
                    
                    items = df_ex[df_ex['Tarima'] == t_orig]
                    for _, item in items.iterrows():
                        st.session_state.BD_Detalle_Tarimas = pd.concat([st.session_state.BD_Detalle_Tarimas, pd.DataFrame([{"ID_Detalle": len(st.session_state.BD_Detalle_Tarimas)+1, "ID_Tarima": nuevo_id_tpm, "SKU": item['Producto/SKU'], "PO": item['PO'], "Cantidad": item['Cantidad']}])], ignore_index=True)
                
                # PERSISTENCIA: Se suben los archivos modificados a GitHub automáticamente
                subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                st.success("¡Plantilla integrada correctamente y respaldada en GitHub!")
            except Exception as e: 
                st.error(f"Error en procesamiento: {e}")
# =============================================================================
# SECCIÓN 7B: MÓDULO TARIMAS - INTERFAZ DE TABLA CON CASILLAS E IMPRESIÓN PDF
# =============================================================================
    if not st.session_state.BD_Tarimas.empty:
        st.write("---")
        st.subheader("🖨️ Panel de Impresión Masiva de Tarimas")
        
        def resaltar_nuevas(row):
            return ['background-color: #FFF59D' if row['Es_Nueva'] else '' for _ in row]
            
        df_estilado = st.session_state.BD_Tarimas.style.apply(resaltar_nuevas, axis=1)
        seleccion_tabla = st.dataframe(df_estilado, use_container_width=True, column_order=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus"], on_select="rerun", selection_mode="multi-row")
        filas_seleccionadas = seleccion_tabla.get("selection", {}).get("rows", [])
        
        if filas_seleccionadas:
            tarimas_elegidas = st.session_state.BD_Tarimas.iloc[filas_seleccionadas]['ID_Tarima'].tolist()
            if len(tarimas_elegidas) == 1:
                t_imp = tarimas_elegidas
                st.download_button(label=f"📥 Descargar PDF Tarima #{t_imp}", data=generar_pdf_tarima(t_imp), file_name=f"Tarima_{t_imp}.pdf", mime="application/pdf")
            else:
                if st.button("📦 Unificar y Preparar Lote de Impresión"):
                    buffer_lote = io.BytesIO()
                    doc_lote = SimpleDocTemplate(buffer_lote, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
                    story_lote, styles = [], getSampleStyleSheet()
                    
                    for t_imp in tarimas_elegidas:
                        detalles = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == t_imp]
                        tarima_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == t_imp].iloc
                        
                        style_g = ParagraphStyle('G_L', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1)
                        story_lote.append(Spacer(1, 1.8 * inch)); story_lote.append(Paragraph(f"TARIMA<br/><br/><b>#{t_imp}</b>", style_g)); story_lote.append(PageBreak())
                        
                        style_n, style_ng = styles['Normal'], ParagraphStyle('NG_L', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)
                        story_lote.append(Paragraph(f"<b>Detalle Interno - Tarima #{t_imp}</b>", styles['Heading2']))
                        story_lote.append(Paragraph(f"<b>Operador:</b> {tarima_info['Creado_Por']} | <b>Fecha:</b> {tarima_info['Fecha_Creacion']}", style_n))
                        story_lote.append(Spacer(1, 0.3 * inch))
                        
                        for _, item in detalles.iterrows():
                            art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
                            nom_art = art.iloc['Nombre'] if not art.empty else "Desconocido"
                            story_lote.append(Paragraph(f"<b>PO:</b> {item['PO']} | <b>SKU:</b> {item['SKU']} - {nom_art}", style_n))
                            story_lote.append(Spacer(1, 0.4 * inch)); story_lote.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", style_ng))
                        story_lote.append(PageBreak())
                        
                    if story_lote: story_lote.pop()
                    doc_lote.build(story_lote, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
                    st.download_button(label="📥 Descargar Lote Completo (PDF)", data=buffer_lote.getvalue(), file_name="Lote_Tarimas.pdf", mime="application/pdf")
        else: 
            st.warning("Seleccione una o más filas en la tabla para habilitar las descargas de PDFs.")
# =============================================================================
# SECCIÓN 7C: MÓDULO REMISIONES - GENERACIÓN Y CONFIRMACIÓN DE ENVÍOS
# =============================================================================
elif opcion_menu == "🚚 Módulo Remisiones":
    st.title("🚚 Generación de Remisiones de Salida")
    t_disp = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible']['ID_Tarima'].tolist()
    
    if not t_disp: 
        st.warning("⚠️ No existen tarimas disponibles bajo estatus 'Disponible' para despacho.")
    else:
        t_sel = st.multiselect("Seleccione Tarimas para este Envío:", options=t_disp)
        col_e, col_r = st.columns(2)
        with col_e:
            nom_e = st.text_input("Líder / Emisor:", "Jesus Morales")
            dir_e = st.text_input("Almacén de Origen:", "Metales")
        with col_r:
            nom_r = st.text_input("Receptor / Cliente:", "Galvatec Industrias")
            dir_r = st.text_input("Dirección Destino:", "Prol. Valle Guadiana 919, Parque Industrial II, Parque Industrial, 35078 Gómez Palacio, Dgo., Mexico")
            
        if not is_admin: 
            st.error("🔒 Operación Bloqueada: Requiere contraseña de Administrador para confirmar la salida.")
        else:
            if st.button("🚀 Confirmar Salida y Generar Nueva Remisión"):
                if not t_sel or not nom_r: 
                    st.error("Complete los campos obligatorios antes de proceder.")
                else:
                    n_id = len(st.session_state.BD_Datos_Generales_Remision) + 1
                    fol = f"E00{26 + n_id}"
                    reg = {"ID_Remision": n_id, "Folio_Remision": fol, "Fecha_Hora_Salida": datetime.datetime.now().strftime("%d/%m/%Y"), "Nombre_Emisor": nom_e, "Direccion_Emisor": dir_e, "Nombre_Receptor": nom_r, "Direccion_Receptor": dir_r, "Tarimas_Asociadas": t_sel}
                    
                    st.session_state.BD_Datos_Generales_Remision = pd.concat([st.session_state.BD_Datos_Generales_Remision, pd.DataFrame([reg])], ignore_index=True)
                    st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(t_sel), 'Estatus'] = 'Remesada'
                    
                    # PERSISTENCIA EN GITHUB: Se guardan las actualizaciones permanentes en el repositorio
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success(f"✅ ¡Remisión {fol} Generada y Guardada de Forma Permanente!")
                    
    if not st.session_state.BD_Datos_Generales_Remision.empty:
        st.write("---")
        st.subheader("🖨️ Descarga Documental de Remisiones")
        r_sel = st.selectbox("Seleccione Folio para Descarga:", st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].unique())
        row = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'] == r_sel].iloc
        df_det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(row['Tarimas_Asociadas'])]
        
        c1, c2 = st.columns(2)
        with c1: st.download_button("📥 Descargar Remisión (PDF)", data=generar_pdf_remision_general(row, df_det), file_name=f"Remision_{r_sel}.pdf", mime="application/pdf")
        with c2: st.download_button("📥 Descargar Anexo Tarimas (PDF)", data=generar_pdf_anexo_tarimas(row['Tarimas_Asociadas'], df_det), file_name=f"Anexo_{r_sel}.pdf", mime="application/pdf")
