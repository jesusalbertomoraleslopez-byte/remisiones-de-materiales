import streamlit as st
import os
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
st.set_page_config(
    page_title="Remisiones de Materiales",
    layout="wide",
    page_icon="favicon.png" if os.path.exists("favicon.png") else "📦"
)

# Inyectar CSS de Imagen Corporativa Oficial (Industria SIGRAMA)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;700&family=Questrial&display=swap');

    /* Fuentes globales */
    html, body, [class*="css"], .stApp {
        font-family: 'Questrial', sans-serif !important;
        background-color: #FFFFFF !important;
    }

    h1, h2, h3, h4, h5, h6, .main-title {
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        color: #111111 !important;
    }

    /* Barra lateral corporativa en Negro profundo #111111 */
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
        border-right: 1px solid #1E293B !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p, 
    [data-testid="stSidebar"] span, 
    [data-testid="stSidebar"] label {
        color: #FFFFFF !important;
        font-family: 'Questrial', sans-serif !important;
    }
    
    /* Logo negativo en sidebar */
    [data-testid="stSidebar"] img {
        filter: grayscale(1) invert(1) brightness(1.2) contrast(1.2) !important;
    }

    /* Botones de navegación en barra lateral */
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        color: #FFFFFF !important;
        font-size: 14px !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        color: #EC2024 !important;
    }

    /* Estilos para inputs de contraseña en barra lateral */
    [data-testid="stSidebar"] input {
        background-color: #1E293B !important;
        color: #FFFFFF !important;
        border-color: #334155 !important;
    }
    [data-testid="stSidebar"] input:focus {
        border-color: #EC2024 !important;
    }

    /* Estilo de Botones Oficiales - Rojo Corporativo #EC2024 */
    div.stButton > button,
    div.stDownloadButton > button,
    div.stFormSubmitButton > button,
    button[data-testid="baseButton-secondary"]:not([role="tab"]):not([data-baseweb="tab"]),
    button[data-testid="baseButton-primary"]:not([role="tab"]):not([data-baseweb="tab"]),
    button[kind="secondary"]:not([role="tab"]):not([data-baseweb="tab"]),
    button[kind="primary"]:not([role="tab"]):not([data-baseweb="tab"]) {
        background-color: #EC2024 !important;
        color: #FFFFFF !important;
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        border-radius: 4px !important;
        border: 1px solid #EC2024 !important;
        padding: 8px 20px !important;
        transition: all 0.3s ease !important;
        text-transform: uppercase;
        font-size: 13px !important;
    }
    div.stButton > button:hover,
    div.stDownloadButton > button:hover,
    div.stFormSubmitButton > button:hover,
    button[data-testid="baseButton-secondary"]:not([role="tab"]):not([data-baseweb="tab"]):hover,
    button[data-testid="baseButton-primary"]:not([role="tab"]):not([data-baseweb="tab"]):hover,
    button[kind="secondary"]:not([role="tab"]):not([data-baseweb="tab"]):hover,
    button[kind="primary"]:not([role="tab"]):not([data-baseweb="tab"]):hover {
        background-color: #FFFFFF !important;
        color: #EC2024 !important;
        border: 1px solid #EC2024 !important;
        box-shadow: 0 4px 12px rgba(236, 32, 36, 0.15) !important;
    }

    /* Tarjetas de Métricas */
    [data-testid="metric-container"] {
        background-color: #FFFFFF !important;
        border: 1px solid #D2D3D5 !important;
        border-left: 5px solid #EC2024 !important;
        border-radius: 4px !important;
        padding: 12px 18px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05) !important;
    }
    [data-testid="metric-container"] label {
        font-family: 'Montserrat', sans-serif !important;
        color: #111111 !important;
        font-weight: 500 !important;
    }
    [data-testid="metric-container"] div[data-testid="stMetricValue"] {
        font-family: 'Montserrat', sans-serif !important;
        color: #EC2024 !important;
        font-weight: 700 !important;
    }
    
    /* Configuración del Editor de Datos y Tablas */
    .stTable header, th {
        background-color: #111111 !important;
        color: #FFFFFF !important;
        font-family: 'Montserrat', sans-serif !important;
    }
    
    /* Inputs y Selectores */
    div[data-baseweb="input"], div[data-baseweb="select"], textarea {
        border-color: #D2D3D5 !important;
        border-radius: 4px !important;
    }
    div[data-baseweb="input"]:focus-within, div[data-baseweb="select"]:focus-within {
        border-color: #EC2024 !important;
    }

    /* Estilo de las pestañas */
    button[role="tab"] {
        font-family: 'Montserrat', sans-serif !important;
        font-weight: bold !important;
        color: #111111 !important;
    }
    
    button[role="tab"][aria-selected="true"] {
        color: #EC2024 !important;
        border-bottom-color: #EC2024 !important;
    }

    /* Reset del File Uploader para encajar en el estilo secundario */
    [data-testid="stFileUploader"] button,
    .stFileUploader button {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 1px solid #D2D3D5 !important;
        border-radius: 4px !important;
        padding: 6px 12px !important;
        font-weight: 500 !important;
        box-shadow: none !important;
        transform: none !important;
    }
    
    [data-testid="stFileUploader"] button *,
    .stFileUploader button * {
        background-color: transparent !important;
        color: #111111 !important;
    }
    
    [data-testid="stFileUploader"] button:hover,
    .stFileUploader button:hover {
        background-color: #F8F9FA !important;
        border-color: #EC2024 !important;
        color: #EC2024 !important;
    }
    
    [data-testid="stFileUploader"] button:hover *,
    .stFileUploader button:hover * {
        color: #EC2024 !important;
    }

    .report-card {
        background-color: #FFFFFF;
        border: 1px solid #D2D3D5;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)

# Renderizado de Banner Corporativo Adaptable
try:
    banner_img = Image.open("REMISIONES APP.png")
    st.image(banner_img, use_container_width=True)
except FileNotFoundError:
    st.warning("⚠️ Cargando interfaz gráfica del banner superior corporativo...")

# Slogan de Resultados / Transformación Principal
st.markdown('<p style="text-align: center; font-size: 16px; font-weight: bold; color: #EC2024; font-family: \'Montserrat\', sans-serif; margin-top: 15px; text-transform: uppercase; letter-spacing: 1px;">SOLUCIONES QUE TRANSFORMAN TU EMPRESA</p>', unsafe_allow_html=True)
st.markdown('<hr style="border: 1px solid #EC2024; margin: 15px 0;">', unsafe_allow_html=True)

# =============================================================================
# 2. MOTOR DE PERSISTENCIA CERTIFICADO (CONEXIÓN UNIFICADA POR API GITHUB)
# =============================================================================
REPO_OWNER = "jesusalbertomoraleslopez-byte"
REPO_NAME = "remisiones-de-materiales"
BRANCH = "main"

def cargar_excel_desde_github(file_name):
    """Carga el archivo Excel: primero busca en disco local (modo offline), luego en GitHub API."""
    import os
    # 1. Intentar cargar localmente primero
    if os.path.exists(file_name):
        try:
            try:
                return pd.read_excel(file_name, sheet_name='Datos_Sistema')
            except Exception:
                return pd.read_excel(file_name, sheet_name=0)
        except Exception as e:
            st.warning(f"⚠️ Error al leer archivo local {file_name}: {e}")
            
    # 2. Si no existe localmente, intentar GitHub
    try:
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{file_name}?ref={BRANCH}"
        headers = {}
        if "github_token" in st.secrets:
            headers["Authorization"] = f"token {st.secrets['github_token']}"
            headers["Accept"] = "application/vnd.github.v3+json"
            
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            datos_json = res.json()
            contenido_base64 = datos_json["content"]
            archivo_bytes = base64.b64decode(contenido_base64)
            try:
                return pd.read_excel(io.BytesIO(archivo_bytes), sheet_name='Datos_Sistema')
            except Exception:
                return pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=0)
    except Exception:
        pass
    return None

def subir_excel_a_github(file_name, dataframe_to_save):
    """Guarda el archivo localmente y luego intenta sincronizarlo con GitHub si hay token disponible."""
    # 1. Guardar localmente siempre
    try:
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            dataframe_to_save.to_excel(writer, index=False, sheet_name='Datos_Sistema')
    except Exception as e:
        st.error(f"⚠️ Error al guardar archivo localmente {file_name}: {e}")
        
    # 2. Sincronizar con GitHub si el token está disponible
    if "github_token" not in st.secrets or not st.secrets["github_token"]:
        # Modo local puro sin token, retornamos True ya que se guardó localmente
        return True
        
    try:
        GITHUB_TOKEN = st.secrets["github_token"]
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{file_name}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}

        # Convertir DataFrame a bytes de Excel en memoria
        buffer_git = io.BytesIO()
        with pd.ExcelWriter(buffer_git, engine='openpyxl') as writer:
            dataframe_to_save.to_excel(writer, index=False, sheet_name='Datos_Sistema')

        base64_content = base64.b64encode(buffer_git.getvalue()).decode("utf-8")
        
        # Obtener el SHA del archivo existente para poder reemplazarlo
        res_get = requests.get(url, headers=headers)
        sha = res_get.json().get("sha") if res_get.status_code == 200 else None

        payload = {
            "message": f"Sincronizacion App: {file_name}", 
            "content": base64_content, 
            "branch": BRANCH
        }
        if sha: 
            payload["sha"] = sha

        res_put = requests.put(url, json=payload, headers=headers)
        return res_put.status_code in [200, 201]

    except Exception as e:
        st.warning(f"⚠️ No se pudo sincronizar con GitHub: {e}")
        return False

@st.cache_data(ttl=60)
def obtener_skus_con_imagen():
    """Obtiene el conjunto de SKUs que tienen una imagen asociada local o remotamente."""
    import os
    import requests
    skus = set()
    
    # 1. Escaneo local
    if os.path.exists("imagenes_articulos"):
        for f in os.listdir("imagenes_articulos"):
            if "(" in f:
                sku = f.split("(")[0]
                skus.add(sku)
                
    # 2. Escaneo remoto (GitHub)
    if "github_token" in st.secrets and st.secrets["github_token"]:
        try:
            GITHUB_TOKEN = st.secrets["github_token"]
            url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
            headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
            res = requests.get(url_list, headers=headers)
            if res.status_code == 200:
                for item in res.json():
                    if "(" in item["name"]:
                        sku = item["name"].split("(")[0]
                        skus.add(sku)
        except Exception:
            pass
            
    return skus

def renderizar_explorador_imagenes():
    st.write("##### 🖼️ Carpeta de Imágenes de Artículos")
    st.markdown("Visualice las imágenes guardadas en el sistema y descargue la base de datos de imágenes completa en formato ZIP:")

    import os
    import zipfile
    import io
    import requests

    # Botón de sincronización con GitHub
    if st.button("🔄 Sincronizar Imágenes con GitHub", use_container_width=True, key="btn_sync_images_explorer"):
        if "github_token" in st.secrets and st.secrets["github_token"]:
            try:
                GITHUB_TOKEN = st.secrets["github_token"]
                url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
                headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
                res_list = requests.get(url_list, headers=headers)
                if res_list.status_code == 200:
                    items_git = res_list.json()
                    downloaded_count = 0
                    os.makedirs("imagenes_articulos", exist_ok=True)
                    for it in items_git:
                        git_file_path = f"imagenes_articulos/{it['name']}"
                        if not os.path.exists(git_file_path):
                            if descargar_imagen_desde_github(git_file_path):
                                downloaded_count += 1
                    st.success(f"✅ Sincronización completada. Se descargaron {downloaded_count} imágenes nuevas de GitHub.")
                    st.rerun()
                else:
                    st.error(f"Error al listar repositorio: Código de estado {res_list.status_code}")
            except Exception as e_sync:
                st.error(f"Error al sincronizar: {e_sync}")
        else:
            st.warning("⚠️ Token de GitHub no configurado para sincronizar de forma remota.")

    st.write("---")

    # Generación de ZIP en memoria
    buf_zip = io.BytesIO()
    has_images = False
    if os.path.exists("imagenes_articulos"):
        file_list_local = [f for f in os.listdir("imagenes_articulos") if os.path.isfile(os.path.join("imagenes_articulos", f)) and not f.startswith(".")]
        if len(file_list_local) > 0:
            has_images = True
            with zipfile.ZipFile(buf_zip, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for file in file_list_local:
                    file_path = os.path.join("imagenes_articulos", file)
                    zip_file.write(file_path, file)
    buf_zip.seek(0)

    col_explorer1, col_explorer2 = st.columns([3, 1])
    with col_explorer2:
        st.download_button(
            label="📥 Descargar Todo (ZIP)",
            data=buf_zip.getvalue(),
            file_name="imagenes_articulos.zip",
            mime="application/zip",
            disabled=not has_images,
            use_container_width=True,
            key="btn_download_images_zip"
        )
        
    with col_explorer1:
        cant_imagenes = len(os.listdir("imagenes_articulos")) if os.path.exists("imagenes_articulos") else 0
        st.write(f"📁 **Directorio:** `imagenes_articulos/` ({cant_imagenes} archivos locales)")

    # Listar archivos
    if has_images:
        files_sorted = sorted([f for f in os.listdir("imagenes_articulos") if os.path.isfile(os.path.join("imagenes_articulos", f)) and not f.startswith(".")])
        for f in files_sorted:
            f_path = os.path.join("imagenes_articulos", f)
            f_size = os.path.getsize(f_path) / 1024.0 # KB
            
            c_img, c_info, c_actions = st.columns([1, 3, 1])
            with c_img:
                try:
                    st.image(f_path, width=60)
                except Exception:
                    st.write("🖼️") # Fallback si la imagen está corrupta
            with c_info:
                st.markdown(f"📄 **Archivo:** `{f}`")
                st.write(f"⚖️ **Tamaño:** `{f_size:.1f} KB`")
            with c_actions:
                try:
                    with open(f_path, "rb") as file_bytes:
                        st.download_button(
                            label="📥 Descargar",
                            data=file_bytes.read(),
                            file_name=f,
                            mime="image/png" if f.endswith(".png") else "image/jpeg",
                            key=f"btn_dl_single_img_{f}"
                        )
                except Exception:
                    st.write("Error")
            st.write("---")
    else:
        st.info("Actualmente no hay imágenes descargadas localmente en la carpeta `imagenes_articulos/`. Haz clic en 'Sincronizar Imágenes con GitHub' para descargar las imágenes que existan en el repositorio.")

def generar_pdf_catalogo_articulos(df_articulos):
    import io
    import os
    import glob
    import requests
    import datetime
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    
    buf = io.BytesIO()
    
    # 36pt margins -> 540pt width
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    
    story = []
    
    # Decoración de fondo / Cabecera repetitiva
    def draw_catalog_decorations(canvas, doc):
        canvas.saveState()
        # Franja roja en el encabezado
        canvas.setFillColor(colors.HexColor("#EC2024"))
        canvas.rect(0, 750, 612, 42, fill=True, stroke=False)
        
        # Logotipo (si existe)
        logo_path = "logo_sigrama.png"
        if os.path.exists(logo_path):
            canvas.drawImage(logo_path, 36, 755, width=120, height=32, mask='auto')
            
        # Título en la franja roja
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 12)
        canvas.drawString(180, 765, "REPORTE DE CATALOGO MAESTRO DE ARTICULOS")
        
        # Fecha de generación en la esquina
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(576, 765, datetime.date.today().strftime("%d-%b-%Y"))
        
        # Pie de página
        canvas.setStrokeColor(colors.HexColor("#EC2024"))
        canvas.setLineWidth(1)
        canvas.line(36, 50, 576, 50)
        
        canvas.setFillColor(colors.HexColor("#616161"))
        canvas.setFont("Helvetica", 8)
        canvas.drawString(36, 38, "Industria SIGRAMA S.A. de C.V. - Sistema Remisiones")
        canvas.drawRightString(576, 38, f"Página {canvas._pageNumber}")
        canvas.restoreState()

    styles = getSampleStyleSheet()
    
    # Crear estilos personalizados si no existen
    try:
        style_header = ParagraphStyle(
            'CatalogHeaderStyle',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=10,
            textColor=colors.white,
            alignment=1 # Centrado
        )
    except Exception:
        style_header = styles['Normal']
        
    try:
        style_cell = ParagraphStyle(
            'CatalogCellText',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=9,
            textColor=colors.HexColor("#212121"),
            leading=11
        )
    except Exception:
        style_cell = styles['Normal']

    try:
        style_cell_bold = ParagraphStyle(
            'CatalogCellTextBold',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=9,
            textColor=colors.HexColor("#212121"),
            leading=11
        )
    except Exception:
        style_cell_bold = styles['Normal']

    # Título en el body
    story.append(Spacer(1, 10))
    
    # Construcción de la tabla
    # Columnas: SKU, Imagen, Descripción, Calibre, Dimensiones, Acabado
    tabla_datos = [[
        Paragraph("SKU", style_header),
        Paragraph("Imagen", style_header),
        Paragraph("Descripción", style_header),
        Paragraph("Calibre", style_header),
        Paragraph("Dimensiones", style_header),
        Paragraph("Acabado", style_header)
    ]]
    
    # Pre-cargar lista remota si es necesario para evitar múltiples llamadas en bucle
    github_items = []
    if "github_token" in st.secrets and st.secrets["github_token"]:
        try:
            GITHUB_TOKEN = st.secrets["github_token"]
            url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
            headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
            res_list = requests.get(url_list, headers=headers)
            if res_list.status_code == 200:
                github_items = res_list.json()
        except Exception:
            pass

    for _, row in df_articulos.iterrows():
        sku = str(row['SKU']).strip()
        nombre = str(row['Nombre'])
        calibre = str(row['Calibre_Espesor'])
        dimensiones = str(row['Dimensiones_Pieza'])
        acabado = str(row['Acabado_Superficial'])
        
        # Buscar imagen
        img_path = None
        matching_local = glob.glob(f"imagenes_articulos/{sku}(*.*")
        if matching_local:
            img_path = matching_local[0]
        else:
            # Buscar en lista de GitHub pre-cargada
            for git_item in github_items:
                if git_item["name"].startswith(f"{sku}("):
                    github_file_path = f"imagenes_articulos/{git_item['name']}"
                    if descargar_imagen_desde_github(github_file_path):
                        img_path = github_file_path
                        break
        
        # Celda de imagen
        img_cell = ""
        if img_path and os.path.exists(img_path):
            try:
                # Usar tamaño 60x60
                img_cell = RLImage(img_path, width=60, height=60, hAlign='CENTER')
            except Exception:
                img_cell = Paragraph("<font color='red'>Error de carga</font>", style_cell)
        else:
            img_cell = Paragraph("<font color='#9e9e9e'>Sin imagen</font>", style_cell)
            
        tabla_datos.append([
            Paragraph(f"<b>{sku}</b>", style_cell_bold),
            img_cell,
            Paragraph(nombre, style_cell),
            Paragraph(calibre, style_cell),
            Paragraph(dimensiones, style_cell),
            Paragraph(acabado, style_cell)
        ])
        
    # Ancho total: 540 pt
    t_cat = Table(tabla_datos, colWidths=[95, 80, 125, 75, 95, 70])
    t_cat.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#EC2024")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E0E0E0")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,1), (1,-1), 'CENTER'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6)
    ]))
    
    story.append(t_cat)
    
    doc.build(story, onFirstPage=draw_catalog_decorations, onLaterPages=draw_catalog_decorations)
    buf.seek(0)
    return buf.getvalue()


def subir_imagen_a_github(file_path):
    """Sube una imagen local a GitHub utilizando API REST, codificando la ruta de forma segura."""
    import os
    import urllib.parse
    if not os.path.exists(file_path):
        return False
    if "github_token" not in st.secrets or not st.secrets["github_token"]:
        return True
    try:
        with open(file_path, "rb") as f:
            base64_content = base64.b64encode(f.read()).decode("utf-8")
        
        GITHUB_TOKEN = st.secrets["github_token"]
        quoted_path = urllib.parse.quote(file_path.replace("\\", "/"))
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{quoted_path}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
        
        res_get = requests.get(url, headers=headers)
        sha = res_get.json().get("sha") if res_get.status_code == 200 else None
        
        payload = {
            "message": f"Sincronizacion de Imagen: {file_path}",
            "content": base64_content,
            "branch": BRANCH
        }
        if sha:
            payload["sha"] = sha
            
        res_put = requests.put(url, json=payload, headers=headers)
        return res_put.status_code in [200, 201]
    except Exception as e:
        st.warning(f"⚠️ No se pudo sincronizar la imagen {file_path} con GitHub: {e}")
        return False


def descargar_imagen_desde_github(file_path):
    """Intenta descargar la imagen de GitHub si no existe en local."""
    import os
    import urllib.parse
    if os.path.exists(file_path):
        return True
    if "github_token" not in st.secrets or not st.secrets["github_token"]:
        return False
    try:
        quoted_path = urllib.parse.quote(file_path.replace("\\", "/"))
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{quoted_path}?ref={BRANCH}"
        headers = {"Authorization": f"token {st.secrets['github_token']}", "Accept": "application/vnd.github.v3+json"}
        
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            datos_json = res.json()
            contenido_base64 = datos_json["content"]
            with open(file_path, "wb") as f:
                f.write(base64.b64decode(contenido_base64))
            return True
    except Exception:
        pass
    return False


def eliminar_imagen_de_github(file_path):
    """Elimina la imagen localmente y en el repositorio de GitHub."""
    import os
    import urllib.parse
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
        except Exception as e:
            st.error(f"⚠️ Error al borrar imagen local {file_path}: {e}")
            
    if "github_token" not in st.secrets or not st.secrets["github_token"]:
        return True
    try:
        GITHUB_TOKEN = st.secrets["github_token"]
        quoted_path = urllib.parse.quote(file_path.replace("\\", "/"))
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{quoted_path}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
        
        res_get = requests.get(url, headers=headers)
        if res_get.status_code == 200:
            sha = res_get.json().get("sha")
            payload = {
                "message": f"Eliminar Imagen: {file_path}",
                "sha": sha,
                "branch": BRANCH
            }
            res_delete = requests.delete(url, json=payload, headers=headers)
            return res_delete.status_code in [200, 204]
    except Exception as e:
        st.warning(f"⚠️ No se pudo eliminar la imagen {file_path} de GitHub: {e}")
    return False


# =============================================================================
# 3. CAPA DE INICIALIZACIÓN GLOBAL SECTORIZADA (BLINDAJE DE SEGURIDAD)
# =============================================================================

# --- Estado de Inicio de Sesión ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "rol" not in st.session_state:
    st.session_state.rol = None
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = None

# --- Catálogo Maestro de Artículos (Actualizado para Persistencia en GitHub) ---
if "BD_Articulos" not in st.session_state or st.session_state.get("BD_Articulos") is None:
    df_git_articulos = cargar_excel_desde_github("BD_Articulos.xlsx")
    if df_git_articulos is not None:
        if "SKU" in df_git_articulos.columns:
            df_git_articulos["SKU"] = df_git_articulos["SKU"].astype(str).str.strip().str.upper()
        st.session_state.BD_Articulos = df_git_articulos
    else:
        # Estructura oficial estricta del sistema en caso de que el archivo no exista aún en GitHub
        st.session_state.BD_Articulos = pd.DataFrame(columns=["SKU", "Nombre", "Calibre_Espesor", "Dimensiones_Pieza", "Acabado_Superficial"])


# --- Catálogo Maestro de Tarimas ---
if "BD_Tarimas" not in st.session_state:
    df_git_tarimas = cargar_excel_desde_github("BD_Tarimas.xlsx")
    if df_git_tarimas is not None:
        st.session_state.BD_Tarimas = df_git_tarimas
    else:
        st.session_state.BD_Tarimas = pd.DataFrame(columns=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus", "Es_Nueva"])

# --- Detalle Granular de Contenido por Tarima ---
if "BD_Detalle_Tarimas" not in st.session_state or st.session_state.get("BD_Detalle_Tarimas") is None:
    df_git_detalles = cargar_excel_desde_github("BD_Detalle_Tarimas.xlsx")
    if df_git_detalles is not None:
        if "SKU" in df_git_detalles.columns:
            df_git_detalles["SKU"] = df_git_detalles["SKU"].astype(str).str.strip().str.upper()
        st.session_state.BD_Detalle_Tarimas = df_git_detalles
    else:
        st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Proyecto", "Parcialidad", "Descripcion", "Cantidad"])

# --- Datos Históricos de Remisiones Oficiales ---
if "BD_Datos_Generales_Remision" not in st.session_state:
    df_git_remisiones = cargar_excel_desde_github("BD_Datos_Generales_Remision.xlsx")
    if df_git_remisiones is not None:
        if "Tarimas_Asociadas" in df_git_remisiones.columns:
            df_git_remisiones["Tarimas_Asociadas"] = df_git_remisiones["Tarimas_Asociadas"].astype(str)
        st.session_state.BD_Datos_Generales_Remision = df_git_remisiones
    else:
        st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=["ID_Remision", "Folio_Remision", "Fecha_Hora_Salida", "Nombre_Emisor", "Direccion_Emisor", "Nombre_Receptor", "Direccion_Receptor", "Tarimas_Asociadas"])

def sincronizar_estatus_tarimas(auto_save=True):
    """Compara las tarimas remesadas con las remisiones activas y corrige huérfanas."""
    if "BD_Tarimas" not in st.session_state or "BD_Datos_Generales_Remision" not in st.session_state:
        return 0
    
    df_tarimas = st.session_state.BD_Tarimas
    df_remisiones = st.session_state.BD_Datos_Generales_Remision
    
    if df_tarimas.empty:
        return 0
        
    import ast
    remesadas_activas = set()
    if not df_remisiones.empty:
        for _, row in df_remisiones.iterrows():
            val = row.get('Tarimas_Asociadas')
            if isinstance(val, str):
                try:
                    t_list = ast.literal_eval(val)
                except Exception:
                    t_list = [val]
            elif isinstance(val, list):
                t_list = val
            else:
                t_list = []
            for t in t_list:
                remesadas_activas.add(str(t).strip())
                
    corregidas = 0
    for idx, row in df_tarimas.iterrows():
        t_id = str(row['ID_Tarima']).strip()
        est = str(row['Estatus']).strip()
        
        # Si está marcada como Remesada pero no está en ninguna remisión activa, vuelve a Disponible
        if est == 'Remesada' and t_id not in remesadas_activas:
            df_tarimas.at[idx, 'Estatus'] = 'Disponible'
            corregidas += 1
        # Si está marcada como Disponible pero sí está en una remisión activa, se corrige a Remesada
        elif est == 'Disponible' and t_id in remesadas_activas:
            df_tarimas.at[idx, 'Estatus'] = 'Remesada'
            corregidas += 1
            
    if corregidas > 0:
        st.session_state.BD_Tarimas = df_tarimas
        if auto_save:
            subir_excel_a_github("BD_Tarimas.xlsx", df_tarimas)
            
    return corregidas

# --- Corrección automática de Estatus de Tarimas ---
if "BD_Tarimas" in st.session_state and "BD_Datos_Generales_Remision" in st.session_state:
    try:
        sincronizar_estatus_tarimas(auto_save=True)
    except Exception:
        pass

# --- Catálogo Operativo de Líderes ---
if "BD_Lideres" not in st.session_state:
    df_git_lideres = cargar_excel_desde_github("BD_Lideres.xlsx")
    if df_git_lideres is not None:
        st.session_state.BD_Lideres = df_git_lideres
    else:
        st.session_state.BD_Lideres = pd.DataFrame([{"ID_Lider": "LID-01", "Nombre_Lider": "Jesus Morales", "Area": "Metales", "Estatus": "Activo"}])

# --- Función Auxiliar y Consecutivo Dinámico de Tarimas (TPM) ---
def obtener_siguiente_consecutivo_tpm():
    # 1. Intentar cargar el consecutivo manual configurado desde disco local
    import os
    if os.path.exists("consecutivo_override.txt"):
        try:
            with open("consecutivo_override.txt", "r", encoding="utf-8") as f:
                val = f.read().strip()
                if val.isdigit():
                    return int(val)
        except Exception:
            pass

    # 2. De lo contrario, calcular dinámicamente a partir de la base de datos
    if "BD_Tarimas" not in st.session_state or st.session_state.BD_Tarimas.empty:
        return 1
    try:
        # Extraer números de los IDs tipo TPM-XXXX
        ids = st.session_state.BD_Tarimas['ID_Tarima'].astype(str)
        numeros = []
        for id_val in ids:
            partes = id_val.split('-')
            if len(partes) > 1 and partes[1].strip().isdigit():
                numeros.append(int(partes[1].strip()))
            elif id_val.startswith('TPM') and id_val[3:].strip().isdigit():
                numeros.append(int(id_val[3:].strip()))
        if numeros:
            return max(numeros) + 1
    except Exception:
        pass
    return len(st.session_state.BD_Tarimas) + 1

if "siguiente_numero_tpm" not in st.session_state or st.session_state["siguiente_numero_tpm"] is None:
    st.session_state["siguiente_numero_tpm"] = obtener_siguiente_consecutivo_tpm()





# =============================================================================
# CAPA DE CONTROL DOCUMENTAL Y DISEÑO IMPRESO CORPORATIVO (INDUSTRIA SIGRAMA)
# =============================================================================
def draw_sigrama_reporte_decorations(canvas, doc):
    """Dibuja los elementos de marca institucionales y diseño de encabezado corporativo del Reporte Consolidado."""
    canvas.saveState()
    # Franja superior roja Sigrama
    canvas.setFillColor(colors.HexColor("#D32F2F"))
    canvas.rect(36, 745, 540, 4, fill=1, stroke=0)
    
    # Marcador de Control de Calidad y Metadatos de Revisión - Eliminados por requerimiento del cliente
    
    # --- CAMBIO AQUÍ: Fecha más grande, destacada en negrita y desplazada hacia abajo ---
    canvas.setFont("Helvetica-Bold", 10) 
    canvas.drawString(36, 730, datetime.date.today().strftime("%d de %B %Y"))
    
    # Título Central del Formato Oficial
    canvas.setFont("Helvetica-Bold", 13)
    canvas.drawCentredString(285, 755, "REPORTE CONSOLIDADO DE INVENTARIO POR FILTRO")
    
    # Pie de Página Legal y Control del SGC (FO-SGC-02) - Eliminado por requerimiento del cliente
    canvas.restoreState()

def generar_pdf_reporte_filtrado(filtros_dict, df_resultado_piezas):
    """Construye el documento PDF oficial de Reporte Consolidado de Inventario con el panel de filtros y la cuadrícula."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    # Estilos de texto especializados para tablas y párrafos
    style_blanco_bold = ParagraphStyle('BB_Rep', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=9)
    style_normal_bold = ParagraphStyle('NB_Rep', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8)
    style_normal_text = ParagraphStyle('NT_Rep', parent=styles['Normal'], fontSize=8)
    
    story.append(Spacer(1, 0.1 * inch))
    
    # --- PANEL LOGÍSTICO DE FILTROS APLICADOS ---
    t_header_filtros = Table([[Paragraph("CRITERIOS DE FILTRADO ACTIVO EN CONSULTA", style_blanco_bold)]], colWidths=[7.5 * inch])
    t_header_filtros.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#757575")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header_filtros)
    
    datos_panel_filtros = [
        [Paragraph("ORDEN DE COMPRA (PO):", style_normal_bold), Paragraph(str(filtros_dict['PO']), style_normal_text),
         Paragraph("SKU / PRODUCTO:", style_normal_bold), Paragraph(str(filtros_dict['SKU']), style_normal_text)],
        [Paragraph("PROYECTO INTERNO:", style_normal_bold), Paragraph(str(filtros_dict['Proyecto']), style_normal_text),
         Paragraph("ID TARIMA:", style_normal_bold), Paragraph(str(filtros_dict['Tarima']), style_normal_text)],
        [Paragraph("PARCIALIDAD:", style_normal_bold), Paragraph(str(filtros_dict['Parcialidad']), style_normal_text),
         Paragraph("ESTATUS ENVÍO:", style_normal_bold), Paragraph(str(filtros_dict['Estatus']), style_normal_text)],
        [Paragraph("DESCRIPCIÓN DE PROYECTO:", style_normal_bold), Paragraph(str(filtros_dict['Descripcion']), style_normal_text),
         Paragraph("TOTAL PIEZAS EN SELECCIÓN:", style_normal_bold), Paragraph(f"<b>{int(df_resultado_piezas['Cantidad'].sum()):,} PZS</b>", style_normal_bold)]
    ]
    
    t_panel = Table(datos_panel_filtros, colWidths=[2.0 * inch, 1.75 * inch, 2.0 * inch, 1.75 * inch])
    t_panel.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F5F5F5"))
    ]))
    story.append(t_panel)
    story.append(Spacer(1, 0.2 * inch))
    
    # --- CUADRÍCULA FORMAL DE MATERIALES FILTRADOS ---
    tabla_materiales = [[
        Paragraph("ID TARIMA", style_blanco_bold),
        Paragraph("OC (PO)", style_blanco_bold),
        Paragraph("PROYECTO", style_blanco_bold),
        Paragraph("SKU / PRODUCTO", style_blanco_bold),
        Paragraph("CANTIDAD", style_blanco_bold),
        Paragraph("ESTATUS", style_blanco_bold)
    ]]
    
    # Mapeo de tarimas a su información de remisión para enriquecer el Estatus
    mapa_remisiones = {}
    if "BD_Datos_Generales_Remision" in st.session_state and not st.session_state.BD_Datos_Generales_Remision.empty:
        import ast
        for _, rem_row in st.session_state.BD_Datos_Generales_Remision.iterrows():
            t_val = rem_row['Tarimas_Asociadas']
            receptor = rem_row.get('Nombre_Receptor', 'Desconocido')
            fecha_salida = rem_row.get('Fecha_Hora_Salida', '')
            t_list = []
            if isinstance(t_val, str):
                try:
                    t_list = ast.literal_eval(t_val)
                except Exception:
                    t_list = [t_val]
            elif isinstance(t_val, list):
                t_list = t_val
                
            for t in t_list:
                mapa_remisiones[str(t).strip()] = {"Receptor": receptor, "Fecha": fecha_salida}
    
    for _, row in df_resultado_piezas.iterrows():
    
        sku_actual = row.get('SKU', '')
        descripcion_final = "Articulo No Registrado en BD Remisiones"
    
        # --- SOLUCIÓN DE BUG: Inicializar variables para evitar UnboundLocalError ---
        calibre = ""
        dims = ""
        acabado = ""
        nombre_com = ""
    
        if "BD_Articulos" in st.session_state and not st.session_state.BD_Articulos.empty:
            df_match = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == sku_actual]
            if not df_match.empty:
                art_info = df_match.iloc[0]
                nombre_com = str(art_info.get('Nombre', '')).strip()
                calibre = str(art_info.get('Calibre_Espesor', '')).strip()
                dims = str(art_info.get('Dimensiones_Pieza', '')).strip()
                acabado = str(art_info.get('Acabado_Superficial', '')).strip()
    
                detalles = []
                if calibre and calibre.lower() != 'nan' and calibre != '': detalles.append(f"<b>Calibre/Espesor:</b> {calibre}")
                if dims and dims.lower() != 'nan' and dims != '': detalles.append(f"<b>Dimensiones:</b> {dims}")
                if acabado and acabado.lower() != 'nan' and acabado != '': detalles.append(f"<b>Material/Acabado:</b> {acabado}")
    
                espec_str = f" | ".join(detalles)
                if espec_str:
                    descripcion_final = f"<b>{nombre_com}</b><br/><font color='#555555' size='7.5'>{espec_str}</font>"
                else:
                    descripcion_final = f"<b>{nombre_com}</b>"
            else:
                descripcion_final = "Articulo No Registrado en BD Remisiones"
        else:
            descripcion_final = "Articulo No Registrado en BD Remisiones"





        
        
        # Buscar si existe una imagen cargada para este SKU
        import glob
        img_encontrada = None
        matching_imgs = glob.glob(f"imagenes_articulos/{sku_actual}(*.*")
        if matching_imgs:
            img_encontrada = matching_imgs[0]
        else:
            if "github_token" in st.secrets and st.secrets["github_token"]:
                try:
                    GITHUB_TOKEN = st.secrets["github_token"]
                    url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
                    headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
                    res_list = requests.get(url_list, headers=headers)
                    if res_list.status_code == 200:
                        items_git = res_list.json()
                        for it in items_git:
                            if it["name"].startswith(f"{sku_actual}("):
                                github_file_path = f"imagenes_articulos/{it['name']}"
                                if descargar_imagen_desde_github(github_file_path):
                                    img_encontrada = github_file_path
                                    break
                except Exception:
                    pass

        desc_paragraph = Paragraph(f"{row['SKU']}<br/><font color='#616161'>{descripcion_final}</font>", style_normal_text)
        if img_encontrada and os.path.exists(img_encontrada):
            from reportlab.platypus import Image as RLImage
            img_flowable = RLImage(img_encontrada, width=75, height=75, hAlign='LEFT')
            sub_t = Table([[img_flowable, desc_paragraph]], colWidths=[80, 2.3 * inch - 80])
            sub_t.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0)
            ]))
            desc_cell_flowables = sub_t
        else:
            desc_cell_flowables = desc_paragraph

        # Generar texto de estatus extendido
        estatus_texto = str(row['Estatus_Envio'])
        if estatus_texto.lower() == "remesado":
            t_id = str(row['ID_Tarima']).strip()
            if t_id in mapa_remisiones:
                receptor_info = str(mapa_remisiones[t_id]["Receptor"])
                fecha_info = str(mapa_remisiones[t_id]["Fecha"]).split()[0] if mapa_remisiones[t_id]["Fecha"] else ""
                # Si el receptor es muy largo, cortarlo para que no desborde la celda
                if len(receptor_info) > 18:
                    receptor_info = receptor_info[:15] + "..."
                estatus_texto = f"<b>Remesado</b><br/><font color='#555555' size='7'>{receptor_info}<br/>{fecha_info}</font>"

        tabla_materiales.append([
            Paragraph(str(row['ID_Tarima']), style_normal_text),
            Paragraph(str(row['PO']), style_normal_text),
            Paragraph(str(row['Proyecto']), style_normal_text),
            desc_cell_flowables,
            Paragraph(f"<b>{int(row['Cantidad'])}</b> Pzs", style_normal_text),
            Paragraph(estatus_texto, style_normal_text)
        ])
        
    t_mat = Table(tabla_materiales, colWidths=[1.0 * inch, 1.1 * inch, 1.2 * inch, 2.3 * inch, 0.9 * inch, 1.0 * inch])
    t_mat.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_mat)
    
    # CORRECTO
    doc.build(story, onFirstPage=draw_sigrama_reporte_decorations, onLaterPages=draw_sigrama_reporte_decorations)


    buffer.seek(0)
    return buffer



# =============================================================================
# DECORADORES BASE CORPORATIVOS PARA DOCUMENTOS DE EMBARQUE
# =============================================================================

def draw_sigrama_decorations(canvas, doc):
    """Dibuja los elementos de marca institucionales y diseño de encabezado corporativo del documento de Embarque."""
    canvas.saveState()
    # Franja superior roja Sigrama
    canvas.setFillColor(colors.HexColor("#D32F2F"))
    canvas.rect(36, 745, 540, 4, fill=1, stroke=0)
    
    # Marcador de Control de Calidad y Metadatos de Revisión - Eliminados por requerimiento del cliente
    
    # --- CAMBIO SOLICITADO: Fecha más grande (font 10) y desplazada hacia abajo (coordenada 730) ---
    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawString(36, 730, datetime.date.today().strftime("%d de %B %Y"))
    
    # Título Central del Formato Oficial
    canvas.setFont("Helvetica-Bold", 13)
    canvas.drawCentredString(285, 755, "EMBARQUE-RECEPCIÓN DE MERCANCÍA")
    
    # Pie de Página Legal y Control del SGC (FO-SGC-02) - Eliminado por requerimiento del cliente
    canvas.restoreState()


# =============================================================================
# FUNCIÓN DE REMISIÓN GENERAL (EMBARQUE-RECEPCIÓN DE MERCANCÍA)
# =============================================================================
def generar_pdf_remision_general(datos_remision, df_detalles_remision):
    """Construye el documento oficial de despacho rellenando las páginas de ReportLab con datos válidos."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=95, bottomMargin=60)
    story, styles = [], getSampleStyleSheet()
    
    # Estilos tipográficos especializados
    style_blanco_bold = ParagraphStyle('WB_Rem', parent=styles['Normal'], textColor=colors.white, fontName="Helvetica-Bold", alignment=1, fontSize=9)
    style_normal_bold = ParagraphStyle('NB_Rem', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8)
    style_normal_text = ParagraphStyle('NT_Rem', parent=styles['Normal'], fontSize=8)
    
    story.append(Spacer(1, 0.1 * inch))
    
    # --- PANEL 1: ENCABEZADO LOGÍSTICO OPERATIVO ---
    t_header = Table([[Paragraph("DATOS GENERALES DE EMBARQUE Y DESPACHO", style_blanco_bold)]], colWidths=[7.5 * inch])
    t_header.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#757575")), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_header)
    
    datos_panel = [
        [Paragraph("FOLIO REMISIÓN:", style_normal_bold), Paragraph(f"<b>{str(datos_remision['Folio_Remision'])}</b>", style_normal_bold),
         Paragraph("FECHA EMISIÓN:", style_normal_bold), Paragraph(str(datos_remision['Fecha_Hora_Salida']), style_normal_text)],
        [Paragraph("EMISOR / ALMACÉN:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Emisor']), style_normal_text),
         Paragraph("ORIGEN:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Emisor']), style_normal_text)],
        [Paragraph("RECEPTOR / CLIENTE:", style_normal_bold), Paragraph(str(datos_remision['Nombre_Receptor']), style_normal_text),
         Paragraph("DESTINO PLANTA:", style_normal_bold), Paragraph(str(datos_remision['Direccion_Receptor']), style_normal_text)]
    ]
    t_panel = Table(datos_panel, colWidths=[1.8 * inch, 1.95 * inch, 1.8 * inch, 1.95 * inch])
    t_panel.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#F5F5F5"))
    ]))
    story.append(t_panel)
    story.append(Spacer(1, 0.2 * inch))
    
    # --- PANEL 2: TABLA DE MATERIALES ASOCIADOS EN DESPACHO ---
    tabla_materiales = [[
        Paragraph("ID TARIMA", style_blanco_bold),
        Paragraph("ORDEN COMPRA", style_blanco_bold),
        Paragraph("PROYECTO", style_blanco_bold),
        Paragraph("SKU / PRODUCTO", style_blanco_bold),
        Paragraph("CANTIDAD", style_blanco_bold)
    ]]
    
    for _, row in df_detalles_remision.iterrows():
        # Búsqueda segura del nombre comercial del artículo
        sku_partida = row.get('SKU', '')
        concepto_remision = "Material de Embarque"
        
        if "BD_Articulos" in st.session_state and not st.session_state.BD_Articulos.empty:
            df_match_rem = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == sku_partida]
            if not df_match_rem.empty:
                art_info = df_match_rem.iloc[0]
                nombre_com = str(art_info.get('Nombre', '')).strip()
                calibre = str(art_info.get('Calibre_Espesor', '')).strip()
                dims = str(art_info.get('Dimensiones_Pieza', '')).strip()
                acabado = str(art_info.get('Acabado_Superficial', '')).strip()
                
                detalles = []
                if calibre and calibre.lower() != 'nan' and calibre != '': detalles.append(f"<b>Calibre/Espesor:</b> {calibre}")
                if dims and dims.lower() != 'nan' and dims != '': detalles.append(f"<b>Dimensiones:</b> {dims}")
                if acabado and acabado.lower() != 'nan' and acabado != '': detalles.append(f"<b>Material/Acabado:</b> {acabado}")
                
                espec_str = f" | ".join(detalles)
                if espec_str:
                    concepto_remision = f"<b>{nombre_com}</b><br/><font color='#555555' size='7.5'>{espec_str}</font>"
                else:
                    concepto_remision = f"<b>{nombre_com}</b>"
            else:
                concepto_remision = "Articulo No Registrado en BD Remisiones"
        else:
            concepto_remision = "Articulo No Registrado en BD Remisiones"



        
        
        # Buscar si existe una imagen cargada para este SKU
        import glob
        img_encontrada = None
        matching_imgs = glob.glob(f"imagenes_articulos/{sku_partida}(*.*")
        if matching_imgs:
            img_encontrada = matching_imgs[0]
        else:
            if "github_token" in st.secrets and st.secrets["github_token"]:
                try:
                    GITHUB_TOKEN = st.secrets["github_token"]
                    url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
                    headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
                    res_list = requests.get(url_list, headers=headers)
                    if res_list.status_code == 200:
                        items = res_list.json()
                        for item in items:
                            if item["name"].startswith(f"{sku_partida}("):
                                github_file_path = f"imagenes_articulos/{item['name']}"
                                if descargar_imagen_desde_github(github_file_path):
                                    img_encontrada = github_file_path
                                    break
                except Exception:
                    pass

        desc_paragraph = Paragraph(f"{row['SKU']}<br/><font color='#616161'>{concepto_remision}</font>", style_normal_text)
        if img_encontrada and os.path.exists(img_encontrada):
            from reportlab.platypus import Image as RLImage
            img_flowable = RLImage(img_encontrada, width=75, height=75, hAlign='LEFT')
            sub_t = Table([[img_flowable, desc_paragraph]], colWidths=[80, 2.7 * inch - 80])
            sub_t.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0)
            ]))
            desc_cell_flowables = sub_t
        else:
            desc_cell_flowables = desc_paragraph
            
        tabla_materiales.append([
            Paragraph(str(row['ID_Tarima']), style_normal_text),
            Paragraph(str(row['PO']), style_normal_text),
            Paragraph(str(row['Proyecto']), style_normal_text),
            desc_cell_flowables,
            Paragraph(f"<b>{int(row['Cantidad'])}</b> Pzs", style_normal_text)
        ])

        
    t_mat = Table(tabla_materiales, colWidths=[1.2 * inch, 1.3 * inch, 1.3 * inch, 2.7 * inch, 1.0 * inch])
    t_mat.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D32F2F")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#757575")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_mat)
    
    # Compilación forzando la inyección de páginas en el objeto Canvas
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

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    import unicodedata
    texto_norm = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto_norm.strip().lower()

def mostrar_pantalla_login():
    st.markdown("""
    <h2 style="text-align: center; font-family: 'Montserrat', sans-serif; color: #111111; font-weight: 700; font-size: 28px; margin: 10px 0;">
        <span style="color: #EC2024;">🔑</span> Acceso al Sistema
    </h2>
    """, unsafe_allow_html=True)
    
    col_log1, col_log2, col_log3 = st.columns([1, 2, 1])
    with col_log2:
        with st.container(border=True):
            st.markdown('<h3 style="font-family: \'Montserrat\', sans-serif; font-weight: 700; color: #111111; text-align: center; margin-top: 0; font-size: 20px;">Control de Remisiones y Tarimas</h3>', unsafe_allow_html=True)
            st.markdown('<p style="font-family: \'Questrial\', sans-serif; color: #64748B; font-size: 14px; text-align: center; margin-bottom: 20px;">Por favor, ingrese sus credenciales para operar el sistema.</p>', unsafe_allow_html=True)
            
            username_input = st.text_input("Usuario (Nombre de Líder o Admin):", key="login_username")
            password_input = st.text_input("Contraseña:", type="password", key="login_password")
            
            btn_login = st.button("Ingresar", use_container_width=True, key="btn_login_submit")
            
            if btn_login:
                username_norm = normalizar_texto(username_input)
                
                # Obtener contraseña de admin configurada o usar la de respaldo
                admin_pwd = "SigramaMetales2026"
                if "admin_password" in st.secrets:
                    admin_pwd = st.secrets["admin_password"]
                    
                # Verificar si es Administrador con múltiples contraseñas de respaldo para evitar cualquier bloqueo
                contrasenas_permitidas = [admin_pwd, "SigramaMetales2026", "SigramaMetales2025", "Admin2025"]
                if username_norm in ["admin", "administrador"] and password_input in contrasenas_permitidas:
                    st.session_state.logged_in = True
                    st.session_state.rol = "Administrador"
                    st.session_state.usuario_actual = "Administrador"
                    st.success("Sesión iniciada como Administrador.")
                    st.rerun()
                    
                # Verificar si es un Líder Operativo
                lider_encontrado = None
                if "BD_Lideres" in st.session_state and not st.session_state.BD_Lideres.empty:
                    for _, row in st.session_state.BD_Lideres.iterrows():
                        lider_name = row["Nombre_Lider"]
                        if normalizar_texto(lider_name) == username_norm:
                            lider_encontrado = lider_name
                            break
                            
                if lider_encontrado is not None and password_input == "Metales":
                    st.session_state.logged_in = True
                    st.session_state.rol = "Operador"
                    st.session_state.usuario_actual = lider_encontrado
                    st.success(f"Sesión iniciada como {lider_encontrado}.")
                    st.rerun()
                elif username_norm == "operador" and password_input == "Metales":
                    st.session_state.logged_in = True
                    st.session_state.rol = "Operador"
                    st.session_state.usuario_actual = "Operador General"
                    st.success("Sesión iniciada como Operador General.")
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas. Verifique el usuario y la contraseña.")

# Ejecutar control de acceso
if not st.session_state.get("logged_in", False):
    mostrar_pantalla_login()
    st.stop()

# --- INTERFAZ POST-LOGIN ---
if os.path.exists("logo_sigrama.png"):
    st.sidebar.image("logo_sigrama.png", use_container_width=True)

st.sidebar.markdown(f"""
<div style="background-color: #1E293B; border: 1px solid #334155; padding: 12px; border-radius: 6px; margin-bottom: 15px; margin-top: 10px;">
    <p style="margin: 0; color: #FFFFFF; font-family: 'Questrial', sans-serif; font-size: 13px;">
        👤 Usuario: <b>{st.session_state.usuario_actual}</b>
    </p>
    <p style="margin: 5px 0 0 0; color: #EC2024; font-family: 'Montserrat', sans-serif; font-size: 12px; font-weight: bold;">
        🔑 Rol: {st.session_state.rol}
    </p>
</div>
""", unsafe_allow_html=True)

if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True, key="btn_logout_sidebar"):
    st.session_state.logged_in = False
    st.session_state.rol = None
    st.session_state.usuario_actual = None
    st.rerun()

is_admin = (st.session_state.rol == "Administrador")

# --- VALIDACIÓN DE SUPERUSUARIO PARA MANTENIMIENTO ---
st.sidebar.write("---")
st.sidebar.title("🛠️ Área de Soporte")
super_pass_input = st.sidebar.text_input("Contraseña de Soporte / IT:", type="password", key="sidebar_super_pass")
is_super = (super_pass_input == "SigramaMetales2025")

if is_super:
    st.sidebar.success("⚡ Modo Superusuario Activo")

st.sidebar.title("🧭 Navegación")
lista_modulos = [
    "📊 Dashboard e Históricos", 
    "🔍 Centro de Consultas", 
    "📦 Módulo Tarimas", 
    "🚚 Módulo Remisiones",
    "📦 Catálogo de Artículos",
    "🏭 Industria 4.0",
    "📖 Manual de Operación",
    "🏢 Reporte por Receptor",
    "🕰️ Carga Histórica"
]

if is_admin or is_super:
    lista_modulos.append("⚙️ Mantenimiento y Catálogos")

opcion_menu = st.sidebar.radio("Seleccione un Módulo:", lista_modulos)

# Slogan de Resultados en el Cierre de la Barra Lateral
st.sidebar.markdown("""
    <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #D2D3D5;">
        <span style="font-family: 'Questrial', sans-serif; font-style: italic; font-size: 13px; color: #FFFFFF; border-bottom: 2px solid #EC2024; padding-bottom: 4px; display: inline-block;">
            Ingeniería que da resultados!!
        </span>
    </div>
""", unsafe_allow_html=True)

# =============================================================================
# 9. INTERFAZ DE USUARIO: DASHBOARD (CON AUTOREPARACIÓN EN CALIENTE)
# =============================================================================
if opcion_menu == "📊 Dashboard e Históricos":
    st.title("📊 Dashboard Planta Metales Inventario Producto")
    # --- CONTROL DE SEGURIDAD INTERNO ---
    # --- AUTO-REPARACIÓN DE CACHÉ EN LÍNEA 410 ---
    # Si por cookies o caché de Streamlit la variable no se encuentra, la forzamos a existir aquí
    if "BD_Detalle_Tarimas" not in st.session_state or st.session_state.get("BD_Detalle_Tarimas") is None:
        st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Proyecto", "Parcialidad", "Descripcion", "Cantidad"])
    
    # Ahora la copia se ejecutará de forma 100% segura
    df_maestro_dash = st.session_state.BD_Detalle_Tarimas.copy()

    
    col_g1, col_g2, col_g3, col_g4 = st.columns(4)
    with col_g1:
        opciones_global_proy = ["Todos"] + df_maestro_dash['Proyecto'].dropna().unique().tolist() if not df_maestro_dash.empty and 'Proyecto' in df_maestro_dash.columns else ["Todos"]
        proy_global_sel = st.selectbox("Filtrar por Proyecto Interno:", opciones_global_proy, key="dash_global_proy_select_unique")
        
    with col_g2:
        if proy_global_sel != "Todos":
            df_maestro_dash = df_maestro_dash[df_maestro_dash['Proyecto'] == proy_global_sel]
        opciones_global_desc = ["Todas"] + df_maestro_dash['Descripcion'].dropna().unique().tolist() if not df_maestro_dash.empty and 'Descripcion' in df_maestro_dash.columns else ["Todas"]
        desc_global_sel = st.selectbox("Filtrar por Descripción de Proyecto:", opciones_global_desc, key="dash_global_desc_select_unique")

    with col_g3:
        if desc_global_sel != "Todas":
            df_maestro_dash = df_maestro_dash[df_maestro_dash['Descripcion'] == desc_global_sel]
        opciones_global_po = ["Todos"] + df_maestro_dash['PO'].dropna().unique().tolist() if not df_maestro_dash.empty and 'PO' in df_maestro_dash.columns else ["Todos"]
        po_global_sel = st.selectbox("Filtrar por Orden de Compra (PO):", opciones_global_po, key="dash_global_po_select_unique")

    with col_g4:
        est_global_sel = st.selectbox("Filtrar por Estatus de Envío:", ["Todos", "Remesado", "No Remesado"], key="dash_global_est_select_unique")

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

    if po_global_sel != "Todos":
        df_detalles_filtrados = df_detalles_filtrados[df_detalles_filtrados['PO'] == po_global_sel]
        tarimas_validas_f = df_detalles_filtrados['ID_Tarima'].unique()
        df_tarimas_filtradas = df_tarimas_filtradas[df_tarimas_filtradas['ID_Tarima'].isin(tarimas_validas_f)]

    if est_global_sel != "Todos":
        estatus_buscado = "Remesada" if est_global_sel == "Remesado" else "Disponible"
        df_tarimas_filtradas = df_tarimas_filtradas[df_tarimas_filtradas['Estatus'] == estatus_buscado]
        tarimas_validas_f = df_tarimas_filtradas['ID_Tarima'].unique()
        df_detalles_filtrados = df_detalles_filtrados[df_detalles_filtrados['ID_Tarima'].isin(tarimas_validas_f)]

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
        if po_global_sel != "Todos": df_completo = df_completo[df_completo['PO'] == po_global_sel]
        if est_global_sel != "Todos":
            estatus_buscado = "Remesada" if est_global_sel == "Remesado" else "Disponible"
            df_completo = df_completo[df_completo['Estatus'] == estatus_buscado]
            
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
            # Normalizar tipos para evitar error pyarrow.lib.ArrowInvalid
            for col_r in ["Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", "Descripción de Proyecto Planta Rio", "% Avance Salida"]:
                if col_r in resumen_avanzado.columns:
                    resumen_avanzado[col_r] = resumen_avanzado[col_r].astype(str)
            for col_r in ["Cant. Tarimas", "Total Piezas", "Piezas Disponibles", "Piezas Remesadas"]:
                if col_r in resumen_avanzado.columns:
                    resumen_avanzado[col_r] = pd.to_numeric(resumen_avanzado[col_r], errors='coerce').fillna(0).astype(int)
            st.dataframe(resumen_avanzado, use_container_width=True, hide_index=True)
            
            # Botón de Descarga Excel para Cumplimiento y Parcialidades
            from datetime import datetime
            buf_parc = io.BytesIO()
            writer_p = pd.ExcelWriter(buf_parc, engine='openpyxl')
            resumen_avanzado.to_excel(writer_p, index=False, sheet_name='Cumplimiento_Parcialidades')
            
            # --- APLICACIÓN DE DISEÑO CORPORATIVO ---
            from openpyxl.styles import Font, PatternFill, Alignment
            workbook_p = writer_p.book
            header_fill_p = PatternFill(start_color="D32F2F", end_color="D32F2F", fill_type="solid")
            header_font_p = Font(color="FFFFFF", bold=True)
            center_alignment_p = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            sheet_p = workbook_p['Cumplimiento_Parcialidades']
            for cell in sheet_p[1]:
                cell.fill = header_fill_p
                cell.font = header_font_p
                cell.alignment = center_alignment_p
                
            for row in sheet_p.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = center_alignment_p
                    
            for col in sheet_p.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                sheet_p.column_dimensions[column].width = min(max_length + 2, 60)
                
            sheet_p.auto_filter.ref = sheet_p.dimensions
            sheet_p.freeze_panes = "A2"
            
            writer_p.close()
            buf_parc.seek(0)
            
            st.download_button(
                label="📥 Generar Reporte de Parcialidades en Excel (.xlsx)",
                data=buf_parc.getvalue(),
                file_name=f"TAR_Cumplimiento_Parcialidades_{datetime.now().strftime('%Y%m%d:%I:%M%p').lower()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_parcialidades_excel"
            )
        else:
            st.info("No existen parcialidades registradas para el filtro seleccionado.")
    else:
        st.info("No hay datos cargados para segmentar por proyectos.")
        
    st.write("---")
    st.subheader("📋 Estado Detallado del Inventario Entarimado (Maestro)")
    if not st.session_state.BD_Tarimas.empty:
        vista_dash = st.session_state.BD_Tarimas.drop(columns=["Es_Nueva"], errors="ignore").copy()
        # Normalizar tipos para evitar error pyarrow.lib.ArrowInvalid
        for col_d in vista_dash.columns:
            if vista_dash[col_d].dtype == object:
                vista_dash[col_d] = vista_dash[col_d].astype(str)
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
        df_maestro_piezas = pd.merge(st.session_state.BD_Detalle_Tarimas, 
                                     st.session_state.BD_Tarimas[['ID_Tarima', 'Estatus', 'Fecha_Creacion']], 
                                     on='ID_Tarima', how='left')
        df_maestro_piezas['Estatus_Envio'] = df_maestro_piezas['Estatus'].apply(lambda x: "Remesado" if x == "Remesada" else "No Remesado")
        
        # Cuadrícula interactiva de 4 columnas en paralelo
        c_filt1, c_filt2, c_filt3, c_filt4 = st.columns(4)
        with c_filt1:
            opc_po = ["Todos"] + df_maestro_piezas['PO'].dropna().unique().tolist()
            f_po = st.selectbox("Orden de Compra (PO):", opc_po, key="query_po_filter_u")
            opc_sku = ["Todos"] + df_maestro_piezas['SKU'].dropna().unique().tolist()
            f_sku = st.selectbox("SKU / Producto:", opc_sku, key="query_sku_filter_u")
        with c_filt2:
            opc_proy = ["Todos"] + df_maestro_piezas['Proyecto'].dropna().unique().tolist()
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
            
            # Enriquecer con Receptor y Fecha de Envío
            mapa_remisiones_ui = {}
            if "BD_Datos_Generales_Remision" in st.session_state and not st.session_state.BD_Datos_Generales_Remision.empty:
                import ast
                for _, rem_row in st.session_state.BD_Datos_Generales_Remision.iterrows():
                    t_val = rem_row['Tarimas_Asociadas']
                    receptor = rem_row.get('Nombre_Receptor', 'N/A')
                    fecha_salida = rem_row.get('Fecha_Hora_Salida', 'N/A')
                    t_list = []
                    if isinstance(t_val, str):
                        try:
                            t_list = ast.literal_eval(t_val)
                        except Exception:
                            t_list = [t_val]
                    elif isinstance(t_val, list):
                        t_list = t_val
                    for t in t_list:
                        mapa_remisiones_ui[str(t).strip()] = {"Receptor": receptor, "Fecha_Salida": str(fecha_salida).split()[0] if fecha_salida else "N/A"}
            
            df_rep['Enviado a (Receptor)'] = "N/A"
            df_rep['Fecha de Envío'] = "N/A"
            for idx, row in df_rep.iterrows():
                if row['Estatus_Envio'] == 'Remesado':
                    t_id = str(row['ID_Tarima']).strip()
                    info = mapa_remisiones_ui.get(t_id, {})
                    df_rep.at[idx, 'Enviado a (Receptor)'] = info.get("Receptor", "N/A")
                    df_rep.at[idx, 'Fecha de Envío'] = info.get("Fecha_Salida", "N/A")
            
            df_rep = df_rep[["ID_Tarima", "PO", "Proyecto", "Parcialidad", "Descripcion", "SKU", "Cantidad", "Estatus_Envio", "Enviado a (Receptor)", "Fecha de Envío", "Fecha_Creacion"]]
            df_rep.columns = ["ID Tarima", "Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", "Descripción de Proyecto Planta Rio", "SKU / Producto", "Cantidad (Pzs)", "Estatus de Envío", "Enviado a (Receptor)", "Fecha de Envío", "Fecha de Ingreso"]
            
            # Normalizar tipos para evitar error pyarrow.lib.ArrowInvalid al serializar DataFrame mixto
            for col in ["ID Tarima", "Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", "Descripción de Proyecto Planta Rio", "SKU / Producto", "Estatus de Envío", "Enviado a (Receptor)", "Fecha de Envío", "Fecha de Ingreso"]:
                if col in df_rep.columns:
                    df_rep[col] = df_rep[col].astype(str)
            df_rep["Cantidad (Pzs)"] = pd.to_numeric(df_rep["Cantidad (Pzs)"], errors='coerce').fillna(0).astype(int)
            
            total_piezas_consulta = int(df_rep['Cantidad (Pzs)'].sum())
            st.metric("🔢 Total Piezas en Selección:", f"{total_piezas_consulta:,} PZS")
            st.dataframe(df_rep, use_container_width=True, hide_index=True)
            
            # --- PANEL DE ACCIONES DE DESCARGA MULTIFORMATO ---
            st.write("### 📥 Descarga Documental Certificada")
            btn_col1, btn_col2 = st.columns(2)
            
            with btn_col1:
                # Diccionario compilado de filtros para inyectar en el reporte estático
                filtros_aplicados = {
                    "PO": f_po, "SKU": f_sku, "Proyecto": f_proy, 
                    "Tarima": f_tar, "Parcialidad": f_parc, 
                    "Descripcion": f_desc, "Estatus": f_est
                }
                pdf_data = generar_pdf_reporte_filtrado(filtros_aplicados, df_resultado)
                from datetime import datetime
                
                # Generar fecha y hora actual en el formato exacto requerido
                fecha_hora_str = datetime.now().strftime("%Y%m%d:%I:%M%p").lower()
                
                st.download_button(
                    label="📄 Descargar Reporte Oficial en PDF",
                    data=pdf_data,
                    file_name=f"TAR_Lote_de_Tarimas_Separado_{fecha_hora_str}.pdf",
                    mime="application/pdf"
                )

            
            with btn_col2:
                # Construcción del reporte de auditoría multi-hoja con openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment
                from openpyxl.utils import get_column_letter

                # =============================================================================
                # SOLUCIÓN DEFINITIVA: REPORTE DE EXCEL BLINDADO CONTRA DATOS VACÍOS
                # =============================================================================
                buf_c = io.BytesIO()
        
                # 1. Creamos los metadatos básicos del reporte
                df_metadatos = pd.DataFrame([
                    {"Concepto": "DOCUMENTO", "Valor": "REPORTE CONSOLIDADO DE INVENTARIO"},
                    {"Concepto": "EMPRESA", "Valor": "INDUSTRIA SIGRAMA S.A. DE C.V."},
                    {"Concepto": "FECHA DE GENERACIÓN", "Valor": datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
                    {"Concepto": "FILTRO: ORDEN DE COMPRA (PO)", "Valor": str(f_po)},
                    {"Concepto": "FILTRO: PROYECTO INTERNO", "Valor": str(f_proy)},
                    {"Concepto": "FILTRO: PARCIALIDAD", "Valor": str(f_parc)},
                    {"Concepto": "FILTRO: DESCRIPCIÓN PROYECTO", "Valor": str(f_desc)},
                    {"Concepto": "FILTRO: SKU / PRODUCTO", "Valor": str(f_sku)},
                    {"Concepto": "FILTRO: ID TARIMA", "Valor": str(f_tar)},
                    {"Concepto": "FILTRO: ESTATUS DE ENVÍO", "Valor": str(f_est)},
                    {"Concepto": "TOTAL PIEZAS EN SELECCIÓN", "Valor": int(total_piezas_consulta)}
                ])
        
                # 2. Validamos si hay datos reales en la tabla para exportar
                if 'df_rep' in locals() and not df_rep.empty:
                    df_exportar_inventario = df_rep.copy()
                else:
                    # Si no hay datos, creamos una tabla vacía segura con los encabezados oficiales
                    df_exportar_inventario = pd.DataFrame(columns=["ID Tarima", "Orden de Compra (PO)", "Proyecto Interno", "Parcialidad", "Descripción de Proyecto Planta Rio", "SKU / Producto", "Cantidad (Pzs)", "Estatus de Envío", "Enviado a (Receptor)", "Fecha de Envío", "Fecha de Ingreso"])
                    df_metadatos.loc[len(df_metadatos)] = {"Concepto": "AVISO", "Valor": "No se encontraron registros con los filtros seleccionados."}
        
                # 3. Escritura segura y directa en el archivo Excel
                writer_c = pd.ExcelWriter(buf_c, engine='openpyxl')
                df_metadatos.to_excel(writer_c, index=False, sheet_name='Resumen_Filtros')
                df_exportar_inventario.to_excel(writer_c, index=False, sheet_name='Listado_Inventario')
                
                # --- APLICACIÓN DE DISEÑO CORPORATIVO AL EXCEL ---
                workbook = writer_c.book
                
                # Definir estilos
                header_fill = PatternFill(start_color="111111", end_color="111111", fill_type="solid") # Negro Corporativo para contraste o EC2024
                # Vamos a usar el gris oscuro del PDF o Rojo
                header_fill = PatternFill(start_color="D32F2F", end_color="D32F2F", fill_type="solid") # Rojo Institucional
                header_font = Font(color="FFFFFF", bold=True)
                center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                
                for sheet_name in workbook.sheetnames:
                    sheet = workbook[sheet_name]
                    
                    # Formatear encabezados
                    for cell in sheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = center_alignment
                        
                    # Centrar todos los datos
                    for row in sheet.iter_rows(min_row=2):
                        for cell in row:
                            cell.alignment = center_alignment
                            
                    # Ajuste automático del ancho de las columnas
                    for col in sheet.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        # Darle un poco más de margen
                        adjusted_width = (max_length + 2)
                        sheet.column_dimensions[column].width = min(adjusted_width, 60) # Límite máximo de ancho
                        
                    # Configurar filtros y renglón fijo para Listado_Inventario
                    if sheet_name == 'Listado_Inventario':
                        sheet.auto_filter.ref = sheet.dimensions
                        sheet.freeze_panes = "A2"

                writer_c.close()
                
                buf_c.seek(0)


                
                
                st.download_button(
                    label="📥 Generar Reporte de Inventario en Excel (.xlsx)",
                    data=buf_c.getvalue(),
                    file_name=f"TAR_Lote_de_Tarimas_Separado_{datetime.now().strftime('%Y%m%d:%I:%M%p').lower()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_download_consulta_piezas_excel_final_master"
                )
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
                    df_ex["Producto/SKU"] = df_ex["Producto/SKU"].astype(str).str.strip().str.upper()
                    if not st.session_state.BD_Tarimas.empty: st.session_state.BD_Tarimas["Es_Nueva"] = False
                    for t_orig in df_ex['Tarima'].unique():
                        # 1. Leer el consecutivo manual configurado, si no existe usa el conteo base
                        if "siguiente_numero_tpm" not in st.session_state or st.session_state["siguiente_numero_tpm"] is None:
                            st.session_state["siguiente_numero_tpm"] = obtener_siguiente_consecutivo_tpm()
            
                        num_actual = st.session_state["siguiente_numero_tpm"]
                        nuevo_id_tpm = f"TPM-{num_actual:04d}"  #<-- AQUÍ USA TU CONTADOR MANUAL (Ejemplo: TPM-0056)
                        
                        # 2. Registramos los datos de la nueva tarima
                        n_t = {"ID_Tarima": nuevo_id_tpm, "Tarima_Origen_Excel": t_orig, "Fecha_Creacion": datetime.datetime.now().strftime("%d/%m/%Y"), "Ubicacion_Actual": "Metales", "Creado_Por": oper, "Tipo_Tarima": tipo_t, "Estatus": "Disponible", "Es_Nueva": True}
                        st.session_state.BD_Tarimas = pd.concat([st.session_state.BD_Tarimas, pd.DataFrame([n_t])], ignore_index=True)
                        
                        # 3. Incrementamos el contador en +1 para la siguiente tarima de la lista
                        st.session_state["siguiente_numero_tpm"] += 1

                        items = df_ex[df_ex['Tarima'] == t_orig]
                        for _, item in items.iterrows():
                            st.session_state.BD_Detalle_Tarimas = pd.concat([st.session_state.BD_Detalle_Tarimas, pd.DataFrame([{"ID_Detalle": len(st.session_state.BD_Detalle_Tarimas) + 1, "ID_Tarima": nuevo_id_tpm, "SKU": item['Producto/SKU'], "PO": item['PO'], "Proyecto": item['Proyecto'], "Parcialidad": item['Parcialidad'], "Descripcion": item['Descripcion'], "Cantidad": item['Cantidad']}])], ignore_index=True)
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                    # Eliminamos el archivo de override si existe, ya que ha sido consumido
                    import os
                    if os.path.exists("consecutivo_override.txt"):
                        try:
                            os.remove("consecutivo_override.txt")
                        except Exception:
                            pass
                    st.session_state["siguiente_numero_tpm"] = obtener_siguiente_consecutivo_tpm()
                    st.success("¡Inventario respaldado con éxito!"); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
            
    if True:
        st.write("---")
        st.subheader("🖨️ Panel de Impresión Masiva de Tarimas")
        
        # Ordenar las tarimas consecutivamente de manera descendente por ID_Tarima
        df_tarimas_sorted = st.session_state.BD_Tarimas.copy()
        
        def extract_tpm_num(id_val):
            id_str = str(id_val).strip()
            partes = id_str.split('-')
            if len(partes) > 1 and partes[1].strip().isdigit():
                return int(partes[1].strip())
            elif id_str.startswith('TPM') and id_str[3:].strip().isdigit():
                return int(id_str[3:].strip())
            return 0
            
        df_tarimas_sorted['_sort_key'] = df_tarimas_sorted['ID_Tarima'].apply(extract_tpm_num)
        df_tarimas_sorted = df_tarimas_sorted.sort_values(by='_sort_key', ascending=False).drop(columns=['_sort_key']).reset_index(drop=True)
        
        df_estilado = df_tarimas_sorted.style.apply(lambda r: ['background-color: #FFF59D' if r['Es_Nueva'] else '' for _ in r], axis=1)
        seleccion_tabla = st.dataframe(df_estilado, use_container_width=True, column_order=["ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus"], on_select="rerun", selection_mode="multi-row")
        filas_seleccionadas = seleccion_tabla.get("selection", {}).get("rows", [])
        
        # =============================================================================
        # BLOQUE DE IMPRESIÓN MAESTRO DE TARIMAS BLINDADO CONTRA LLAVES HUÉRFANAS
        # =============================================================================
        if filas_seleccionadas:
            elegidas = df_tarimas_sorted.iloc[filas_seleccionadas]['ID_Tarima'].tolist()
            if len(elegidas) == 1:
                id_tarima_limpio = str(elegidas[0]).strip()  # Extracción nativa perfecta
        
                # 1. Intentamos buscar las piezas asociadas a esta tarima en la base de datos
                df_tarima_individual = pd.DataFrame()
                if "BD_Detalle_Tarimas" in st.session_state and not st.session_state.BD_Detalle_Tarimas.empty:
                    # Filtramos asegurando que compare texto puro
                    df_det = st.session_state.BD_Detalle_Tarimas[
                        st.session_state.BD_Detalle_Tarimas['ID_Tarima'].astype(str) == id_tarima_limpio
                    ].copy()
                    if not df_det.empty:
                        # Cruce relacional para anexar Estatus, Fecha_Creacion y Estatus_Envio requeridos por el PDF
                        df_tarima_individual = pd.merge(df_det, 
                                                         st.session_state.BD_Tarimas[['ID_Tarima', 'Estatus', 'Fecha_Creacion']], 
                                                         on='ID_Tarima', how='left')
                        df_tarima_individual['Estatus_Envio'] = df_tarima_individual['Estatus'].apply(lambda x: "Remesado" if x == "Remesada" else "No Remesado")
        
                # 2. VALIDACIÓN CRÍTICA: ¿Tiene artículos esta tarima para imprimir?
                if df_tarima_individual.empty or "SKU" not in df_tarima_individual.columns:
                    # Si no hay piezas asociadas, mostramos un mensaje limpio en lugar de lanzar la pantalla de error
                    st.warning(f"⚠️ La tarima **{id_tarima_limpio}** no contiene artículos o piezas asignadas en la base de datos de detalles. No se puede generar un reporte vacío.")
                else:
                    # Si sí tiene piezas con SKU válido, procedemos a compilar de forma segura
                    filtros_simulados = {
                        "PO": "Todos",
                        "SKU": "Todos",
                        "Proyecto": "Todos",
                        "Tarima": id_tarima_limpio,
                        "Parcialidad": "Todos",
                        "Estatus": "Todos",
                        "Descripcion": "Todos"
                    }
                    
                    try:
                        # Función auxiliar interna para generar la etiqueta de 2 hojas para esta tarima
                        def generar_pdf_etiqueta_tarima_individual(t_imp):
                            buf = io.BytesIO()
                            doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
                            story_l, styles = [], getSampleStyleSheet()
                            
                            style_tarima_titulo = ParagraphStyle('T_Giga_Ind', parent=styles['Heading1'], fontName="Helvetica-Bold", fontSize=140, alignment=1, leading=150, textColor=colors.HexColor("#212121"))
                            style_sub_titulo = ParagraphStyle('S_Giga_Ind', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=26, alignment=1, textColor=colors.HexColor("#D32F2F"))
                            style_normal_bold = ParagraphStyle('N_Bold_Ind', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=11, leading=14)
                            style_normal_text = ParagraphStyle('N_Text_Ind', parent=styles['Normal'], fontSize=11, leading=14)
                            style_blanco_bold = ParagraphStyle('B_Bold_Ind', parent=styles['Normal'], fontName="Helvetica-Bold", textColor=colors.white, alignment=1, fontSize=10)
                            
                            det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == t_imp]
                            t_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == t_imp]
                            
                            op_nom = t_info.iloc[0]['Creado_Por'] if not t_info.empty else "N/A"
                            fe_cre = t_info.iloc[0]['Fecha_Creacion'] if not t_info.empty else "N/A"
                            
                            # HOJA 1: CARÁTULA DE IDENTIFICACIÓN
                            story_l.append(Spacer(1, 1.2 * inch))
                            story_l.append(Paragraph("TARIMA", style_sub_titulo))
                            story_l.append(Spacer(1, 0.2 * inch))
                            
                            num_limpio = str(t_imp).split('-')[-1] if '-' in str(t_imp) else str(t_imp)
                            story_l.append(Paragraph(f"#{num_limpio}", style_tarima_titulo))
                            story_l.append(Spacer(1, 1.5 * inch))
                            
                            tabla_base = Table([
                                [Paragraph("CÓDIGO DE IDENTIFICACIÓN:", style_normal_bold), Paragraph(str(t_imp), style_normal_text)],
                                [Paragraph("OPERADOR DE PLANTA:", style_normal_bold), Paragraph(str(op_nom), style_normal_text)],
                                [Paragraph("FECHA DE EMISIÓN:", style_normal_bold), Paragraph(str(fe_cre), style_normal_text)]
                            ], colWidths=[2.5 * inch, 5.0 * inch])
                            tabla_base.setStyle(TableStyle([
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.HexColor("#E0E0E0")),
                                ('BOTTOMPADDING', (0,0), (-1,-1), 6)
                            ]))
                            story_l.append(tabla_base)
                            story_l.append(PageBreak())
                            
                            # HOJA 2: DETALLE DE MATERIALES ASOCIADOS
                            story_l.append(Spacer(1, 0.1 * inch))
                            story_l.append(Paragraph(f"<b>DETALLE DE MATERIALES ASOCIADOS - CONTROL #{t_imp}</b>", styles['Heading2']))
                            story_l.append(Spacer(1, 0.2 * inch))
                            
                            tabla_detalles = [[
                                Paragraph("ORDEN (PO)", style_blanco_bold),
                                Paragraph("SKU / PRODUCTO", style_blanco_bold),
                                Paragraph("DESCRIPCIÓN COMERCIAL", style_blanco_bold),
                                Paragraph("CANTIDAD", style_blanco_bold)
                            ]]
                            
                            for _, item in det.iterrows():
                                art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
                                art_nom = art.iloc[0]['Nombre'] if not art.empty else "Articulo No Registrado en BD Remisiones"
                                
                                sku_partida = item['SKU']
                                # Buscar si existe una imagen cargada para este SKU
                                import glob
                                img_encontrada = None
                                matching_imgs = glob.glob(f"imagenes_articulos/{sku_partida}(*.*")
                                if matching_imgs:
                                    img_encontrada = matching_imgs[0]
                                else:
                                    if "github_token" in st.secrets and st.secrets["github_token"]:
                                        try:
                                            GITHUB_TOKEN = st.secrets["github_token"]
                                            url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
                                            headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
                                            res_list = requests.get(url_list, headers=headers)
                                            if res_list.status_code == 200:
                                                items_git = res_list.json()
                                                for it in items_git:
                                                    if it["name"].startswith(f"{sku_partida}("):
                                                        github_file_path = f"imagenes_articulos/{it['name']}"
                                                        if descargar_imagen_desde_github(github_file_path):
                                                            img_encontrada = github_file_path
                                                            break
                                        except Exception:
                                            pass

                                desc_paragraph = Paragraph(str(art_nom), style_normal_text)
                                if img_encontrada and os.path.exists(img_encontrada):
                                    from reportlab.platypus import Image as RLImage
                                    img_flowable = RLImage(img_encontrada, width=75, height=75, hAlign='LEFT')
                                    sub_t = Table([[img_flowable, desc_paragraph]], colWidths=[80, 3.5 * inch - 80])
                                    sub_t.setStyle(TableStyle([
                                        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                        ('LEFTPADDING', (0,0), (-1,-1), 0),
                                        ('RIGHTPADDING', (0,0), (-1,-1), 0),
                                        ('TOPPADDING', (0,0), (-1,-1), 0),
                                        ('BOTTOMPADDING', (0,0), (-1,-1), 0)
                                    ]))
                                    desc_comercial_flowables = sub_t
                                else:
                                    desc_comercial_flowables = desc_paragraph

                                tabla_detalles.append([
                                    Paragraph(str(item['PO']), style_normal_text),
                                    Paragraph(str(item['SKU']), style_normal_text),
                                    desc_comercial_flowables,
                                    Paragraph(f"<b>{int(item['Cantidad'])}</b> PZS", style_normal_bold)
                                ])
                                
                            t_grid = Table(tabla_detalles, colWidths=[1.3 * inch, 1.5 * inch, 3.5 * inch, 1.2 * inch])
                            t_grid.setStyle(TableStyle([
                                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")),
                                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
                                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                ('TOPPADDING', (0,0), (-1,-1), 6),
                                ('BOTTOMPADDING', (0,0), (-1,-1), 6)
                            ]))
                            story_l.append(t_grid)
                            
                            doc.build(story_l, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
                            buf.seek(0)
                            return buf

                        # 3. Compilamos el PDF de reporte pasándole sus parámetros obligatorios
                        pdf_datos_compilados = generar_pdf_reporte_filtrado(filtros_simulados, df_tarima_individual)
                        
                        if hasattr(pdf_datos_compilados, 'getvalue'):
                            pdf_bytes_listos = pdf_datos_compilados.getvalue()
                        else:
                            pdf_bytes_listos = pdf_datos_compilados
                            
                        # 4. Generamos el PDF de etiqueta unificada de 2 hojas para esta tarima
                        pdf_etiqueta_bytes = generar_pdf_etiqueta_tarima_individual(id_tarima_limpio).getvalue()
        
                        col_btns1, col_btns2 = st.columns(2)
                        with col_btns1:
                            st.download_button(
                                label=f"📄 Descargar Reporte PDF",
                                data=pdf_bytes_listos,
                                file_name=f"TAR_Reporte_Inventario_{id_tarima_limpio}.pdf",
                                mime="application/pdf",
                                key=f"btn_descarga_maestra_tarima_ind_{id_tarima_limpio}"
                            )
                        with col_btns2:
                            st.download_button(
                                label=f"🖨️ Descargar Etiqueta PDF (2 Hojas)",
                                data=pdf_etiqueta_bytes,
                                file_name=f"TAR_Etiqueta_Identificacion_{id_tarima_limpio}.pdf",
                                mime="application/pdf",
                                key=f"btn_descarga_etiqueta_ind_{id_tarima_limpio}"
                            )
                    except Exception as e:
                        st.error(f"⚠️ Error al procesar el diseño de ReportLab para la tarima {id_tarima_limpio}: {e}")

        



                
            else:

                if st.button("📦 Unificar Lote de Impresión"):
                    buf_1 = io.BytesIO()
                    doc_1 = SimpleDocTemplate(buf_1, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=90, bottomMargin=60)
                    story_l, styles = [], getSampleStyleSheet()
                    
                    # Estilos tipográficos institucionales de gran impacto
                    style_tarima_titulo = ParagraphStyle('T_Giga', parent=styles['Heading1'], fontName="Helvetica-Bold", fontSize=140, alignment=1, leading=150, textColor=colors.HexColor("#212121"))
                    style_sub_titulo = ParagraphStyle('S_Giga', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=26, alignment=1, textColor=colors.HexColor("#D32F2F"))
                    style_normal_bold = ParagraphStyle('N_Bold', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=11, leading=14)
                    style_normal_text = ParagraphStyle('N_Text', parent=styles['Normal'], fontSize=11, leading=14)
                    style_blanco_bold = ParagraphStyle('B_Bold', parent=styles['Normal'], fontName="Helvetica-Bold", textColor=colors.white, alignment=1, fontSize=10)
                    
                    for t_imp in elegidas:
                        det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == t_imp]
                        t_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == t_imp]
                        
                        op_nom = t_info.iloc[0]['Creado_Por'] if not t_info.empty else "N/A"
                        fe_cre = t_info.iloc[0]['Fecha_Creacion'] if not t_info.empty else "N/A"
                        
                        # =============================================================================
                        # HOJA 1: CARÁTULA DE IDENTIFICACIÓN MASIVA (NÚMERO GIGANTE)
                        # =============================================================================
                        story_l.append(Spacer(1, 1.2 * inch))
                        story_l.append(Paragraph("TARIMA", style_sub_titulo))
                        story_l.append(Spacer(1, 0.2 * inch))
                        
                        # Extraemos solo el número limpio (ej. si es TPM-0024 extrae 0024)
                        num_limpio = str(t_imp).split('-')[-1] if '-' in str(t_imp) else str(t_imp)
                        story_l.append(Paragraph(f"#{num_limpio}", style_tarima_titulo))
                        
                        story_l.append(Spacer(1, 1.5 * inch))
                        
                        # Mini panel inferior de trazabilidad en la carátula
                        tabla_base = Table([
                            [Paragraph("CÓDIGO DE IDENTIFICACIÓN:", style_normal_bold), Paragraph(str(t_imp), style_normal_text)],
                            [Paragraph("OPERADOR DE PLANTA:", style_normal_bold), Paragraph(str(op_nom), style_normal_text)],
                            [Paragraph("FECHA DE EMISIÓN:", style_normal_bold), Paragraph(str(fe_cre), style_normal_text)]
                        ], colWidths=[2.5 * inch, 5.0 * inch])
                        tabla_base.setStyle(TableStyle([
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.HexColor("#E0E0E0")),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 6)
                        ]))
                        story_l.append(tabla_base)
                        
                        # Forzamos salto inmediato a la siguiente página
                        story_l.append(PageBreak())
                        
                        # =============================================================================
                        # HOJA 2: DESGLOSE GRANULAR Y DETALLE DE PIEZAS
                        # =============================================================================
                        story_l.append(Spacer(1, 0.1 * inch))
                        story_l.append(Paragraph(f"<b>DETALLE DE MATERIALES ASOCIADOS - CONTROL #{t_imp}</b>", styles['Heading2']))
                        story_l.append(Spacer(1, 0.2 * inch))
                        
                        # Construcción de cuadrícula formal de inventario
                        tabla_detalles = [[
                            Paragraph("ORDEN (PO)", style_blanco_bold),
                            Paragraph("SKU / PRODUCTO", style_blanco_bold),
                            Paragraph("DESCRIPCIÓN COMERCIAL", style_blanco_bold),
                            Paragraph("CANTIDAD", style_blanco_bold)
                        ]]

                    
                        for _, item in det.iterrows():
                            art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
                            if not art.empty:
                                art_info = art.iloc[0]
                                nombre_com = str(art_info.get('Nombre', '')).strip()
                                calibre = str(art_info.get('Calibre_Espesor', '')).strip()
                                dims = str(art_info.get('Dimensiones_Pieza', '')).strip()
                                acabado = str(art_info.get('Acabado_Superficial', '')).strip()
                                
                                detalles = []
                                if calibre and calibre.lower() != 'nan' and calibre != '': detalles.append(f"<b>Calibre/Espesor:</b> {calibre}")
                                if dims and dims.lower() != 'nan' and dims != '': detalles.append(f"<b>Dimensiones:</b> {dims}")
                                if acabado and acabado.lower() != 'nan' and acabado != '': detalles.append(f"<b>Material/Acabado:</b> {acabado}")
                                
                                espec_str = f" | ".join(detalles)
                                if espec_str:
                                    art_nom = f"<b>{nombre_com}</b><br/><font color='#555555' size='8.5'>{espec_str}</font>"
                                else:
                                    art_nom = f"<b>{nombre_com}</b>"
                            else:
                                art_nom = "Articulo No Registrado en BD Remisiones"
                            
                            sku_partida = item['SKU']
                            # Buscar si existe una imagen cargada para este SKU
                            import glob
                            img_encontrada = None
                            matching_imgs = glob.glob(f"imagenes_articulos/{sku_partida}(*.*")
                            if matching_imgs:
                                img_encontrada = matching_imgs[0]
                            else:
                                if "github_token" in st.secrets and st.secrets["github_token"]:
                                    try:
                                        GITHUB_TOKEN = st.secrets["github_token"]
                                        url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
                                        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
                                        res_list = requests.get(url_list, headers=headers)
                                        if res_list.status_code == 200:
                                            items_git = res_list.json()
                                            for it in items_git:
                                                if it["name"].startswith(f"{sku_partida}("):
                                                    github_file_path = f"imagenes_articulos/{it['name']}"
                                                    if descargar_imagen_desde_github(github_file_path):
                                                        img_encontrada = github_file_path
                                                        break
                                    except Exception:
                                        pass

                            desc_paragraph = Paragraph(str(art_nom), style_normal_text)
                            if img_encontrada and os.path.exists(img_encontrada):
                                from reportlab.platypus import Image as RLImage
                                img_flowable = RLImage(img_encontrada, width=75, height=75, hAlign='LEFT')
                                sub_t = Table([[img_flowable, desc_paragraph]], colWidths=[80, 3.5 * inch - 80])
                                sub_t.setStyle(TableStyle([
                                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                                    ('LEFTPADDING', (0,0), (-1,-1), 0),
                                    ('RIGHTPADDING', (0,0), (-1,-1), 0),
                                    ('TOPPADDING', (0,0), (-1,-1), 0),
                                    ('BOTTOMPADDING', (0,0), (-1,-1), 0)
                                ]))
                                desc_comercial_flowables = sub_t
                            else:
                                desc_comercial_flowables = desc_paragraph

                            tabla_detalles.append([
                                Paragraph(str(item['PO']), style_normal_text),
                                Paragraph(str(item['SKU']), style_normal_text),
                                desc_comercial_flowables,
                                Paragraph(f"<b>{int(item['Cantidad'])}</b> PZS", style_normal_bold)
                            ])
                            
                        t_grid = Table(tabla_detalles, colWidths=[1.3 * inch, 1.5 * inch, 3.5 * inch, 1.2 * inch])
                        t_grid.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#757575")),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#BDBDBD")),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('TOPPADDING', (0,0), (-1,-1), 6),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 6)
                        ]))
                        story_l.append(t_grid)
                        
                        # Salto de página para separar la siguiente tarima del lote
                        story_l.append(PageBreak())
                        
                    if story_l: 
                        story_l.pop()  # Remueve el último salto de página sobrante
                        
                    doc_1.build(story_l, onFirstPage=draw_sigrama_decorations, onLaterPages=draw_sigrama_decorations)
                    
                    st.download_button(
                        label="📥 Descargar Lote Completo (PDF)", 
                        data=buf_1.getvalue(), 
                        file_name=f"LOT_Lote_Tarimas_Separadas_{datetime.datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        key="btn_download_lote_tarimas_unificado_final_v2"
                    )


        # =============================================================================
        # NUEVA SECCIÓN: CONSULTA GRANULAR DE ARTÍCULOS POR TARIMA
        # =============================================================================
        st.write("---")
        st.write("### 🔍 Consulta de Contenido de Artículos por Tarima")
        st.info("💡 **Guía de uso:** Seleccione o busque una Tarima específica para desplegar de forma inmediata el listado de materiales y piezas que contiene.")

        # 1. Validar que la base de datos de tarimas tenga registros activos
        if "BD_Tarimas" in st.session_state and not st.session_state.BD_Tarimas.empty:
            
            # Obtener la lista de IDs de tarimas disponibles para el menú desplegable
            lista_tarimas_disponibles = sorted(st.session_state.BD_Tarimas["ID_Tarima"].unique())
            
            # Selector de Tarima para el usuario
            tarima_seleccionada = st.selectbox(
                "📋 Seleccione el ID de la Tarima a consultar:",
                options=lista_tarimas_disponibles,
                index=0,
                key="selector_granular_tarimas_modulo_t"
            )

            if tarima_seleccionada:
                # 2. Filtrar el detalle de materiales asociados a la tarima seleccionada
                if "BD_Detalle_Tarimas" in st.session_state and not st.session_state.BD_Detalle_Tarimas.empty:
                    
                    df_articulos_filtrados = st.session_state.BD_Detalle_Tarimas[
                        st.session_state.BD_Detalle_Tarimas["ID_Tarima"] == tarima_seleccionada
                    ].copy()

                    if not df_articulos_filtrados.empty:
                        st.write(f"##### 📦 Materiales contenidos en la Tarima: **{tarima_seleccionada}**")
                        
                        # Mostrar tabla limpia con datos clave para el operador
                        st.dataframe(
                            df_articulos_filtrados[["SKU", "PO", "Proyecto", "Descripcion", "Cantidad"]],
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        # Resumen de conteo rápido para auditorías
                        total_pzs_tarima = int(df_articulos_filtrados["Cantidad"].sum())
                        st.success(f"📊 Total de piezas contabilizadas en esta tarima: **{total_pzs_tarima} Pzs**")
                    else:
                        st.warning(f"⚠️ La tarima **{tarima_seleccionada}** está registrada pero no cuenta con un desglose de piezas asociado en la base de datos.")
                else:
                    st.error("❌ Error de sistema: No se encuentra cargada la Base de Datos granular de Detalles de Tarimas.")
        else:
            st.warning("⚠️ No existen tarimas registradas en el sistema para realizar consultas de contenido.")








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
                    reg = {"ID_Remision": len(st.session_state.BD_Datos_Generales_Remision)+1, "Folio_Remision": fol, "Fecha_Hora_Salida": datetime.datetime.now().strftime("%d/%m/%Y"), "Nombre_Emisor": nom_e, "Direccion_Emisor": dir_e, "Nombre_Receptor": nom_r, "Direccion_Receptor": dir_r, "Tarimas_Asociadas": str(t_sel)}
                    st.session_state.BD_Datos_Generales_Remision = pd.concat([st.session_state.BD_Datos_Generales_Remision, pd.DataFrame([reg])], ignore_index=True)
                    st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(t_sel), 'Estatus'] = 'Remesada'
                    subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                    subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                    st.success(f"✅ ¡Remisión {fol} Generada y Guardada de Forma Permanente!"); st.rerun()
                    
        # =============================================================================
    # INTERFAZ DE DESCARGA DE REMISIONES CORREGIDA Y GARANTIZADA (FO-MET-10)
    # =============================================================================
    if not st.session_state.BD_Datos_Generales_Remision.empty:
        st.write("---")
        st.subheader("📋 Descarga Documental de Remisiones")
        
        # Filtros de Fechas
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            fecha_inicio = st.date_input("Fecha de Inicio:", value=datetime.date.today() - datetime.timedelta(days=30), key="rem_download_start_date")
        with col_f2:
            fecha_fin = st.date_input("Fecha de Fin:", value=datetime.date.today(), key="rem_download_end_date")
            
        # Parseo de fechas y filtrado
        df_remisiones = st.session_state.BD_Datos_Generales_Remision.copy()
        
        def parsear_fecha_rem(val):
            import datetime as dt
            try:
                return dt.datetime.strptime(str(val).split()[0], "%d/%m/%Y").date()
            except Exception:
                try:
                    return dt.datetime.strptime(str(val).split()[0], "%Y-%m-%d").date()
                except Exception:
                    return dt.date.today()
                    
        df_remisiones['Fecha_Date'] = df_remisiones['Fecha_Hora_Salida'].apply(parsear_fecha_rem)
        df_remisiones_filtradas = df_remisiones[
            (df_remisiones['Fecha_Date'] >= fecha_inicio) & 
            (df_remisiones['Fecha_Date'] <= fecha_fin)
        ]
        
        if df_remisiones_filtradas.empty:
            st.info("ℹ️ No se encontraron remisiones en el rango de fechas seleccionado.")
        else:
            df_mostrar = df_remisiones_filtradas.drop(columns=['Fecha_Date'], errors='ignore')
            
            # Ordenar descendentemente por número de folio
            def extract_folio_num(folio_str):
                import re
                numeros = re.findall(r'\d+', str(folio_str))
                if numeros:
                    return int(numeros[-1])
                return 0
                
            df_mostrar['Folio_Num'] = df_mostrar['Folio_Remision'].apply(extract_folio_num)
            df_mostrar = df_mostrar.sort_values(by='Folio_Num', ascending=False).drop(columns=['Folio_Num'], errors='ignore')
            
            st.write("Seleccione una remisión de la tabla para habilitar la descarga:")
            
            # Normalizar TODOS los tipos de columnas para evitar pyarrow.lib.ArrowInvalid
            for col_m in df_mostrar.columns:
                if df_mostrar[col_m].dtype == object:
                    df_mostrar[col_m] = df_mostrar[col_m].astype(str)
                
            sel_grid = st.dataframe(
                df_mostrar,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row",
                key="tabla_descarga_remisiones_final",
                column_config={
                    "Folio_Remision": st.column_config.TextColumn("Folio de Remisión"),
                    "Fecha_Hora_Salida": st.column_config.TextColumn("Fecha de Salida"),
                    "Nombre_Emisor": st.column_config.TextColumn("Emisor"),
                    "Nombre_Receptor": st.column_config.TextColumn("Receptor"),
                    "Tarimas_Asociadas": st.column_config.TextColumn("Tarimas Asociadas")
                }
            )
            
            filas_seleccionadas = sel_grid.get("selection", {}).get("rows", [])
            if filas_seleccionadas:
                idx_seleccionado = filas_seleccionadas[0]
                row_dict = df_mostrar.iloc[idx_seleccionado].to_dict()
                r_sel = row_dict['Folio_Remision']
                
                # Convertimos las tarimas asociadas de texto a lista real de Python
                import ast
                tarimas_lista = row_dict['Tarimas_Asociadas']
                if isinstance(tarimas_lista, str):
                    try:
                        tarimas_lista = ast.literal_eval(tarimas_lista)
                    except Exception:
                        tarimas_lista = [tarimas_lista]
                
                # Filtramos el desglose granular de las piezas
                df_det = st.session_state.BD_Detalle_Tarimas[
                    st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(tarimas_lista)
                ]
                
                st.write("")
                st.success(f"📌 Folio seleccionado: **{r_sel}**")
                
                # --- RENDERIZADO DE BOTONES EXCLUSIVO ---
                c1, c2 = st.columns(2)
                with c1: 
                    st.download_button(
                        label="📄 Descargar Reporte Oficial", 
                        data=generar_pdf_remision_general(row_dict, df_det), 
                        file_name=f"Remision_{r_sel}.pdf", 
                        key=f"btn_dl_rem_pdf_{r_sel}",
                        mime="application/pdf"
                    )
                with c2: 
                    if not df_det.empty:
                        pass
                    else:
                        st.info("ℹ️ No hay desglose de piezas registrado para el anexo de este folio.")
            else:
                st.info("💡 Seleccione una fila haciendo clic en el extremo izquierdo de la remisión deseada para descargar su reporte.")




# =============================================================================
# SECCIÓN 17A: PANEL DE MANTENIMIENTO AVANZADO Y EDICIÓN EN CALIENTE (SUPERUSER)
# =============================================================================
elif opcion_menu == "📦 Catálogo de Artículos":
    st.title("📦 Catálogo Maestro de Artículos")
    st.markdown("Consulte y filtre de forma dinámica el catálogo oficial de productos cargados en el sistema:")

    if "BD_Articulos" in st.session_state and not st.session_state.BD_Articulos.empty:
        df_articulos_base = st.session_state.BD_Articulos.copy()

        # Cuadrícula de filtros dinámicos (3 columnas en paralelo para optimizar espacio)
        art_col1, art_col2, art_col3 = st.columns(3)
        
        with art_col1:
            # Filtro por SKU
            opc_art_sku = ["Todos"] + sorted(df_articulos_base['SKU'].dropna().unique().tolist())
            f_art_sku = st.selectbox("Filtrar por SKU:", opc_art_sku, key="filter_art_sku")
            
            # Filtro por Dimensiones
            opc_art_dim = ["Todos"] + sorted(df_articulos_base['Dimensiones_Pieza'].dropna().unique().tolist())
            f_art_dim = st.selectbox("Filtrar por Dimensiones:", opc_art_dim, key="filter_art_dim")

        with art_col2:
            # Filtro por Nombre / Descripción Comercial
            opc_art_nom = ["Todos"] + sorted(df_articulos_base['Nombre'].dropna().unique().tolist())
            f_art_nom = st.selectbox("Filtrar por Nombre:", opc_art_nom, key="filter_art_nom")

        with art_col3:
            # Filtro por Calibre / Espesor
            opc_art_cal = ["Todos"] + sorted(df_articulos_base['Calibre_Espesor'].dropna().unique().tolist())
            f_art_cal = st.selectbox("Filtrar por Calibre / Espesor:", opc_art_cal, key="filter_art_cal")
            
            # Filtro por Acabado Superficial
            opc_art_acab = ["Todos"] + sorted(df_articulos_base['Acabado_Superficial'].dropna().unique().tolist())
            f_art_acab = st.selectbox("Filtrar por Acabado Superficial:", opc_art_acab, key="filter_art_acab")

        # Aplicación en cascada de los filtros seleccionados
        df_art_filtrado = df_articulos_base.copy()
        if f_art_sku != "Todos":
            df_art_filtrado = df_art_filtrado[df_art_filtrado['SKU'] == f_art_sku]
        if f_art_nom != "Todos":
            df_art_filtrado = df_art_filtrado[df_art_filtrado['Nombre'] == f_art_nom]
        if f_art_cal != "Todos":
            df_art_filtrado = df_art_filtrado[df_art_filtrado['Calibre_Espesor'] == f_art_cal]
        if f_art_dim != "Todos":
            df_art_filtrado = df_art_filtrado[df_art_filtrado['Dimensiones_Pieza'] == f_art_dim]
        if f_art_acab != "Todos":
            df_art_filtrado = df_art_filtrado[df_art_filtrado['Acabado_Superficial'] == f_art_acab]

        st.write("---")
        
        # Despliegue de métricas rápidas de la consulta
        col_metric, col_pdf = st.columns([2, 1])
        with col_metric:
            st.metric("🔢 Artículos en Selección:", f"{len(df_art_filtrado)} Items")
        with col_pdf:
            st.write("") # Alineación vertical
            pdf_data = generar_pdf_catalogo_articulos(df_art_filtrado)
            st.download_button(
                label="📄 Descargar Catálogo (PDF)",
                data=pdf_data,
                file_name="Reporte_Catalogo_Articulos.pdf",
                mime="application/pdf",
                key="btn_download_catalogo_pdf_tab",
                use_container_width=True
            )
        
        # Despliegue de la tabla de datos estructurada
        st.dataframe(
            df_art_filtrado, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "SKU": st.column_config.TextColumn("SKU / Código"),
                "Nombre": st.column_config.TextColumn("Descripción Comercial"),
                "Calibre_Espesor": st.column_config.TextColumn("Calibre / Espesor"),
                "Dimensiones_Pieza": st.column_config.TextColumn("Dimensiones de la Pieza"),
                "Acabado_Superficial": st.column_config.TextColumn("Acabado Superficial")
            }
        )

        # --- SECCIÓN ADICIONAL: CARGA Y DETALLE DE IMAGEN DE ARTÍCULO ---
        st.write("---")
        st.subheader("🖼️ Detalle e Imagen del Artículo")
        
        filtro_estado_img = st.radio(
            "Filtrar artículos por estado de imagen:",
            ["Todos", "Sin imagen", "Con imagen"],
            horizontal=True,
            key="filtro_estado_imagen_select"
        )
        
        skus_con_img = obtener_skus_con_imagen()
        lista_skus_filtrados = df_art_filtrado['SKU'].dropna().tolist()
        
        if filtro_estado_img == "Sin imagen":
            lista_skus_filtrados = [s for s in lista_skus_filtrados if s not in skus_con_img]
        elif filtro_estado_img == "Con imagen":
            lista_skus_filtrados = [s for s in lista_skus_filtrados if s in skus_con_img]
            
        opc_skus_img = ["Seleccione un SKU..."] + lista_skus_filtrados
        sku_sel = st.selectbox("Seleccione un SKU para administrar su imagen:", opc_skus_img, key="sku_select_img")
        
        if sku_sel != "Seleccione un SKU...":
            art_row = df_art_filtrado[df_art_filtrado['SKU'] == sku_sel].iloc[0]
            
            import glob
            os.makedirs("imagenes_articulos", exist_ok=True)
            
            # Escanear localmente
            matching_files_local = glob.glob(f"imagenes_articulos/{sku_sel}(*.*")
            
            imagen_final_path = None
            if matching_files_local:
                imagen_final_path = matching_files_local[0]
            else:
                # Buscar en GitHub si el token está disponible
                if "github_token" in st.secrets and st.secrets["github_token"]:
                    try:
                        GITHUB_TOKEN = st.secrets["github_token"]
                        url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
                        headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
                        res_list = requests.get(url_list, headers=headers)
                        if res_list.status_code == 200:
                            items = res_list.json()
                            for item in items:
                                if item["name"].startswith(f"{sku_sel}("):
                                    github_file_path = f"imagenes_articulos/{item['name']}"
                                    if descargar_imagen_desde_github(github_file_path):
                                        imagen_final_path = github_file_path
                                        break
                    except Exception:
                        pass
            
            col_ficha, col_cargar = st.columns(2)
            
            with col_ficha:
                st.write("##### Ficha Técnica del Artículo")
                st.markdown(f"**SKU / Código:** `{sku_sel}`")
                st.markdown(f"**Nombre:** {art_row['Nombre']}")
                st.markdown(f"**Calibre / Espesor:** {art_row['Calibre_Espesor']}")
                st.markdown(f"**Dimensiones:** {art_row['Dimensiones_Pieza']}")
                st.markdown(f"**Acabado Superficial:** {art_row['Acabado_Superficial']}")
                
                st.write("")
                if imagen_final_path and os.path.exists(imagen_final_path):
                    st.image(imagen_final_path, caption=f"Imagen cargada para {sku_sel}", use_container_width=True)
                    
                    if st.button("🗑️ Eliminar Imagen de Artículo", use_container_width=True, key=f"btn_del_img_{sku_sel}"):
                        if eliminar_imagen_de_github(imagen_final_path):
                            obtener_skus_con_imagen.clear()
                            st.success("¡Imagen eliminada correctamente del servidor y GitHub!")
                            st.rerun()
                        else:
                            st.error("Error al eliminar la imagen en GitHub.")
                else:
                    st.info("Este artículo no cuenta con una imagen asociada actualmente.")
            
            with col_cargar:
                st.write("##### Cargar / Reemplazar Imagen")
                
                file_uploaded = st.file_uploader("Subir archivo de imagen (PNG, JPG, JPEG):", type=["png", "jpg", "jpeg"], key=f"file_uploader_{sku_sel}")
                
                from streamlit_paste_button import paste_image_button as pbutton
                paste_result = pbutton(
                    "📋 Pegar captura de pantalla desde el portapapeles",
                    key=f"paste_button_{sku_sel}"
                )
                
                nueva_imagen_data = None
                img_ext = ".png"
                
                if file_uploaded:
                    nueva_imagen_data = Image.open(file_uploaded)
                    _, ext = os.path.splitext(file_uploaded.name)
                    if ext.lower() in [".png", ".jpg", ".jpeg"]:
                        img_ext = ext.lower()
                elif paste_result.image_data is not None:
                    nueva_imagen_data = paste_result.image_data
                    img_ext = ".png"
                    
                if nueva_imagen_data is not None:
                    st.write("---")
                    st.warning("⚠️ Vista Previa de la Nueva Imagen (Aún no se ha guardado):")
                    st.image(nueva_imagen_data, caption="Vista Previa de la Carga", use_container_width=True)
                    
                    c_save, c_cancel = st.columns(2)
                    
                    with c_save:
                        if st.button("💾 Guardar y Sincronizar Imagen", use_container_width=True, key=f"btn_save_img_{sku_sel}"):
                            # 1. Crear nombre final en inglés: SKU(dd-Mon-yy)
                            meses_en = {
                                1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
                                7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
                            }
                            hoy = datetime.date.today()
                            fecha_ingles = f"{hoy.day:02d}-{meses_en[hoy.month]}-{str(hoy.year)[-2:]}"
                            nombre_archivo = f"{sku_sel}({fecha_ingles}){img_ext}"
                            nuevo_path = f"imagenes_articulos/{nombre_archivo}"
                            
                            if imagen_final_path and os.path.exists(imagen_final_path):
                                eliminar_imagen_de_github(imagen_final_path)
                            
                            try:
                                if img_ext in [".jpg", ".jpeg"] and nueva_imagen_data.mode in ("RGBA", "P"):
                                    nueva_imagen_data = nueva_imagen_data.convert("RGB")
                                nueva_imagen_data.save(nuevo_path)
                                
                                if subir_imagen_a_github(nuevo_path):
                                    obtener_skus_con_imagen.clear()
                                    st.success("¡Imagen guardada y sincronizada correctamente en GitHub!")
                                    st.rerun()
                                else:
                                    st.error("Error al sincronizar la imagen con el repositorio de GitHub.")
                            except Exception as ex_save:
                                st.error(f"Error al guardar el archivo localmente: {ex_save}")
                                
                    with c_cancel:
                        if st.button("❌ Descartar", use_container_width=True, key=f"btn_discard_img_{sku_sel}"):
                            st.rerun()

    else:
        st.info("ℹ️ No hay artículos registrados en el catálogo maestro actualmente o el archivo en GitHub está vacío.")

elif opcion_menu == "🏭 Industria 4.0":
    st.title("🏭 Manufactura Inteligente e Industria 4.0")
    st.markdown("##### Integración Tecnológica y Automatización de Procesos Logísticos")
    
    st.markdown("""
    <div style="background-color: #F8FAFC; border-left: 5px solid #EC2024; padding: 15px; border-radius: 4px; margin-bottom: 20px;">
        <p style="font-family: 'Questrial', sans-serif; font-size: 15px; color: #1E293B; margin: 0; line-height: 1.6;">
            La digitalización y la interconectividad son los pilares fundamentales del ecosistema industrial contemporáneo. 
            Esta sección describe cómo la presente plataforma de Remisiones de Materiales se alinea estratégicamente 
            con las tendencias de la <b>Manufactura Inteligente</b> y la <b>Industria 4.0</b>.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    tab_ind1, tab_ind2, tab_ind3 = st.tabs([
        "🔬 Justificación de Manufactura Inteligente", 
        "🎯 Beneficios Estratégicos", 
        "💻 Stack Tecnológico"
    ])
    
    with tab_ind1:
        st.write("")
        st.markdown("""
        <h4 style="color: #EC2024; font-family: 'Montserrat', sans-serif; font-weight: bold; margin-bottom: 10px;">
            El Rol del Dato en la Cadena Logística
        </h4>
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6; color: #111111;">
            En la manufactura de metales y ensambles de precisión, la logística interna suele ser un cuello de botella. 
            La transformación digital implica convertir flujos físicos en información estructurada utilizable para la toma de decisiones.
        </p>
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6; color: #111111;">
            <b>Habilitación del Gemelo Digital (Digital Twin):</b> Cada tarima empacada en la planta no es solo un bulto físico de metal; 
            a través de este sistema, se asocia dinámicamente con un identificador único <code>TPM-XXXX</code> que representa digitalmente su 
            contenido granular (SKUs, POs, proyectos y parcialidades). Esto permite rastrear virtualmente la ubicación, 
            estatus e historial de cada paquete sin necesidad de inspecciones físicas manuales, logrando una representación en tiempo real del inventario.
        </p>
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6; color: #111111;">
            <b>Filosofía de IoT Industrial (IIoT):</b> Aunque no utilicemos sensores físicos directos en esta etapa, el esquema de 
            registro descentralizado por medio de terminales móviles y la sincronización con una base de datos centralizada en la nube 
            sigue la arquitectura básica del IIoT: capturar el dato directamente en la fuente (línea de producción / empaque) 
            y comunicarlo de manera ágil hacia las capas de toma de decisiones.
        </p>
        """, unsafe_allow_html=True)
        
    with tab_ind2:
        st.write("")
        st.markdown('<h4 style="color: #EC2024; font-family: \'Montserrat\', sans-serif; font-weight: bold; margin-bottom: 15px;">Impacto y Retorno de Inversión Tecnológica</h4>', unsafe_allow_html=True)
        
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.markdown("""
            <div class="report-card">
                <h5 style="color: #111111; font-family: 'Montserrat', sans-serif; font-weight: bold; margin-top: 0;"><span style="color: #EC2024;">📍</span> Trazabilidad de Extremo a Extremo</h5>
                <p style="font-family: 'Questrial', sans-serif; font-size: 13.5px; line-height: 1.5; color: #334155; margin-bottom: 0;">
                    Registro inalterable de cada movimiento: desde la creación de la tarima por el líder del área, pasando por el empaque de productos específicos, hasta su despacho final documentado en una remisión firmada.
                </p>
            </div>
            <div class="report-card">
                <h5 style="color: #111111; font-family: 'Montserrat', sans-serif; font-weight: bold; margin-top: 0;"><span style="color: #EC2024;">⚡</span> Eliminación del Error Humano</h5>
                <p style="font-family: 'Questrial', sans-serif; font-size: 13.5px; line-height: 1.5; color: #334155; margin-bottom: 0;">
                    El sistema valida automáticamente las cantidades y referencias contra el catálogo maestro de artículos y líderes, impidiendo errores de dedo, duplicados o envío de folios erróneos a los clientes.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
        with col_b2:
            st.markdown("""
            <div class="report-card">
                <h5 style="color: #111111; font-family: 'Montserrat', sans-serif; font-weight: bold; margin-top: 0;"><span style="color: #EC2024;">⏱️</span> Agilidad en Operaciones</h5>
                <p style="font-family: 'Questrial', sans-serif; font-size: 13.5px; line-height: 1.5; color: #334155; margin-bottom: 0;">
                    Generación en 3 segundos de manifiestos y anexos en formato PDF certificado y listos para impresión. Lo que antes requería transcripciones manuales de varias horas, ahora se realiza de forma automatizada e inmediata.
                </p>
            </div>
            <div class="report-card">
                <h5 style="color: #111111; font-family: 'Montserrat', sans-serif; font-weight: bold; margin-top: 0;"><span style="color: #EC2024;">☁️</span> Persistencia Híbrida y Resiliencia</h5>
                <p style="font-family: 'Questrial', sans-serif; font-size: 13.5px; line-height: 1.5; color: #334155; margin-bottom: 0;">
                    El sistema combina almacenamiento local en caché con sincronización instantánea en la nube mediante APIs (GitHub API). Esto garantiza que la aplicación siga operando en planta incluso ante interrupciones de red.
                </p>
            </div>
            """, unsafe_allow_html=True)

    with tab_ind3:
        st.write("")
        st.markdown('<h4 style="color: #EC2024; font-family: \'Montserrat\', sans-serif; font-weight: bold; margin-bottom: 15px;">Arquitectura de Software y Datos</h4>', unsafe_allow_html=True)
        
        st.markdown("""
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6; color: #111111;">
            La aplicación ha sido diseñada bajo estándares de desarrollo ágil y tecnologías robustas de procesamiento de datos:
        </p>
        """, unsafe_allow_html=True)
        
        col_st1, col_st2, col_st3, col_st4 = st.columns(4)
        with col_st1:
            st.metric("🎨 Front-End UI", "Streamlit")
            st.caption("Entorno web reactivo de alta velocidad, estilizado con CSS personalizado de marca.")
        with col_st2:
            st.metric("⚙️ Data Engine", "Python / Pandas")
            st.caption("Procesamiento y modelado de datos tabulares complejos directamente en memoria.")
        with col_st3:
            st.metric("📁 Cloud Storage", "GitHub API")
            st.caption("Mecanismo seguro de control de versiones y almacenamiento distribuido basado en la nube.")
        with col_st4:
            st.metric("📄 Document Maker", "ReportLab")
            st.caption("Motor de renderizado de precisión milimétrica para la construcción de archivos PDF.")

elif opcion_menu == "📖 Manual de Operación":
    st.title("📖 Manual de Operación del Sistema")
    st.markdown("##### Guía de Usuario para el Control de Remisiones y Tarimas")
    
    st.markdown("""
    <div style="background-color: #F8FAFC; border-left: 5px solid #EC2024; padding: 15px; border-radius: 4px; margin-bottom: 20px;">
        <p style="font-family: 'Questrial', sans-serif; font-size: 15px; color: #1E293B; margin: 0; line-height: 1.6;">
            Bienvenido al manual oficial de operación de la aplicación. 
            Utilice las pestañas inferiores para comprender el funcionamiento detallado de cada módulo del sistema.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    tab_man1, tab_man2, tab_man3, tab_man4, tab_man5 = st.tabs([
        "🔑 Acceso y Seguridad", 
        "📦 Módulo Tarimas", 
        "🚚 Módulo Remisiones", 
        "🔍 Consultas e Impresión", 
        "⚙️ Mantenimiento (Admin)"
    ])
    
    with tab_man1:
        st.write("")
        st.markdown("""
        <h4 style="color: #EC2024; font-family: 'Montserrat', sans-serif; font-weight: bold;">Control de Credenciales y Perfiles</h4>
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            La aplicación restringe sus funciones mediante credenciales de acceso para garantizar la integridad de los datos de inventario:
        </p>
        <ul style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            <li><b>Perfil Operador / Líder:</b> Permite operar la mayor parte del sistema (visualizar tableros, cargar proyectos, registrar tarimas, generar remisiones y descargar reportes). 
            Para ingresar, use su nombre tal cual aparece en el catálogo de líderes (ej. <code>Jesús Morales</code>, <code>Miguel Alvarado</code>) y la contraseña general <code>Metales</code>.</li>
            <li><b>Perfil Administrador:</b> Cuenta con privilegios especiales para realizar modificaciones en caliente de bases de datos, corregir inventarios, modificar catálogos y purgar registros. 
            Para ingresar, use el usuario <code>admin</code> y la contraseña corporativa <code>SigramaMetales2026</code>.</li>
            <li><b>Cierre de Sesión:</b> Al concluir sus labores, es recomendable presionar el botón <b>🚪 Cerrar Sesión</b> en la barra lateral para evitar que otros usuarios realicen movimientos bajo su firma.</li>
        </ul>
        """, unsafe_allow_html=True)
        
    with tab_man2:
        st.write("")
        st.markdown("""
        <h4 style="color: #EC2024; font-family: 'Montserrat', sans-serif; font-weight: bold;">Empaque e Inventariado de Tarimas</h4>
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            El Módulo Tarimas sirve para registrar y desglosar el contenido de los paquetes que se preparan en planta:
        </p>
        <ol style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            <li><b>Carga de Plantilla Excel:</b> Suba el archivo en formato Excel con las columnas requeridas (Tarima, Producto/SKU, PO, Proyecto, Parcialidad, Descripcion, Cantidad).</li>
            <li><b>Asignación de Líder y Tipo:</b> Indique el líder de línea a cargo del empaque y especifique si la tarima física es Cuadrada o Rectangular.</li>
            <li><b>Procesamiento:</b> Presione <i>Procesar e Integrar Plantilla Avanzada</i>. El sistema validará los SKUs contra el catálogo oficial, generará automáticamente folios correlativos únicos de tarima (TPM-XXXX), y subirá la información estructurada a las bases de datos de GitHub de manera segura.</li>
        </ol>
        """, unsafe_allow_html=True)
        
    with tab_man3:
        st.write("")
        st.markdown("""
        <h4 style="color: #EC2024; font-family: 'Montserrat', sans-serif; font-weight: bold;">Salida de Materiales y Remisión</h4>
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            El Módulo Remisiones agrupa las tarimas que saldrán de la planta con destino a un cliente:
        </p>
        <ol style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            <li><b>Selección de Tarimas:</b> El dropdown muestra todas las tarimas con estatus <code>Disponible</code> (es decir, creadas pero que aún no han sido asignadas a ninguna remisión). Seleccione una o más tarimas para agruparlas.</li>
            <li><b>Datos de Despacho:</b> Especifique el líder que autoriza la salida, el almacén de origen, el nombre del receptor (cliente) y la dirección destino.</li>
            <li><b>Emisión:</b> Haga clic en <i>🚀 Confirmar Salida y Generar Nueva Remisión</i>. La aplicación generará un folio oficial correlativo (ej. <code>E0028</code>), marcará el estatus de las tarimas como <code>Remesada</code> para que no puedan volver a usarse y guardará la transacción en la nube.</li>
        </ol>
        """, unsafe_allow_html=True)
        
    with tab_man4:
        st.write("")
        st.markdown("""
        <h4 style="color: #EC2024; font-family: 'Montserrat', sans-serif; font-weight: bold;">Centro de Consultas y Descarga Documental</h4>
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            La consulta y obtención de documentos certificados se centraliza en dos áreas clave:
        </p>
        <ul style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            <li><b>Dashboard e Históricos:</b> Muestra métricas agregadas del inventario (tarimas disponibles, remesadas, etc.) y gráficos interactivos de barras y líneas para analizar la cantidad de productos por proyecto en tiempo real.</li>
            <li><b>Centro de Consultas (Impresión Masiva):</b> Permite ver la lista completa de tarimas, filtrarlas por folio, estatus o fecha, y seleccionar filas específicas para generar reportes anejos en PDF con el desglose exacto de su contenido.</li>
            <li><b>Descarga Documental de Remisiones:</b> En el Módulo Remisiones, el dropdown inferior contiene los folios únicos de salida (ordenados de más recientes a más antiguos). Seleccione el folio deseado y presione <b>📥 Descargar Remisión Oficial en PDF</b> para descargar el manifiesto formal con formato de firma e imagen de Industria SIGRAMA.</li>
        </ul>
        """, unsafe_allow_html=True)
        
    with tab_man5:
        st.write("")
        st.markdown("""
        <h4 style="color: #EC2024; font-family: 'Montserrat', sans-serif; font-weight: bold;">Mantenimiento Crítico (Acceso Restringido)</h4>
        <p style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            El módulo <b>⚙️ Mantenimiento y Catálogos</b> está reservado para el perfil <b>Administrador</b> (o personal de soporte de TI):
        </p>
        <ul style="font-family: 'Questrial', sans-serif; font-size: 14px; line-height: 1.6;">
            <li><b>Ajustar Cantidades:</b> Editor interactivo tipo Excel que permite modificar directamente las cantidades o descripciones de una tarima ya guardada para corregir capturas erróneas.</li>
            <li><b>Catálogo de Líderes y Artículos:</b> Permite dar de alta nuevos códigos SKU o empleados autorizados, ya sea uno a uno o cargando un archivo Excel masivo.</li>
            <li><b>Contador de Tarimas (Consecutivos):</b> Si necesita reiniciar la numeración o saltar folios (ej. continuar en 150), digite el consecutivo deseado y presione guardar.</li>
            <li><b>Purga de Datos (Eliminación):</b> Permite eliminar remisiones o tarimas seleccionadas para liberar espacio, o aplicar un <b>Reset de Fábrica Total</b>. 
            <i>Nota: Debido al riesgo operativo, esta pestaña se bloquea por completo a nivel de base de datos si el rol del usuario no es Administrador, incluso si se ingresa con la clave de soporte técnico.</i></li>
        </ul>
        """, unsafe_allow_html=True)

elif opcion_menu == "🏢 Reporte por Receptor":
    st.title("🏢 Reporte de Piezas Remisionadas por Receptor")
    st.markdown("Consulte el detalle de piezas y tarimas enviadas a cada cliente o receptor de remisiones.")
    
    if not st.session_state.BD_Datos_Generales_Remision.empty and not st.session_state.BD_Detalle_Tarimas.empty:
        df_rems = st.session_state.BD_Datos_Generales_Remision.copy()
        receptores_disp = ["Seleccione un Receptor...", "Todos"] + sorted(df_rems['Nombre_Receptor'].dropna().unique().tolist())
        
        col_rec1, col_rec2 = st.columns([3, 1])
        with col_rec1:
            receptor_sel = st.selectbox("Seleccione el Receptor (Cliente):", receptores_disp, key="sel_receptor_reporte")
            
        if receptor_sel != "Seleccione un Receptor...":
            if receptor_sel == "Todos":
                df_rems_filtradas = df_rems.copy()
            else:
                df_rems_filtradas = df_rems[df_rems['Nombre_Receptor'] == receptor_sel]
            
            tarimas_asociadas = []
            import ast
            for _, row in df_rems_filtradas.iterrows():
                t_val = row['Tarimas_Asociadas']
                t_list = []
                if isinstance(t_val, str):
                    try:
                        t_list = ast.literal_eval(t_val)
                    except Exception:
                        t_list = [t_val]
                elif isinstance(t_val, list):
                    t_list = t_val
                
                for t in t_list:
                    tarimas_asociadas.append({
                        "ID_Tarima": str(t).strip(),
                        "Folio_Remision": row['Folio_Remision'],
                        "Fecha_Salida": row['Fecha_Hora_Salida'],
                        "Receptor": row['Nombre_Receptor']
                    })
                    
            if tarimas_asociadas:
                df_tarimas_receptor = pd.DataFrame(tarimas_asociadas)
                df_detalle = st.session_state.BD_Detalle_Tarimas.copy()
                df_detalle['ID_Tarima'] = df_detalle['ID_Tarima'].astype(str).str.strip()
                
                df_cruce = pd.merge(df_tarimas_receptor, df_detalle, on="ID_Tarima", how="inner")
                
                if not df_cruce.empty:
                    if "BD_Articulos" in st.session_state and not st.session_state.BD_Articulos.empty:
                        df_cruce = pd.merge(df_cruce, st.session_state.BD_Articulos[['SKU', 'Nombre']], on="SKU", how="left")
                    else:
                        df_cruce['Nombre'] = "N/A"
                        
                    df_cruce = df_cruce.rename(columns={
                        "ID_Tarima": "Tarima",
                        "Folio_Remision": "Remisión",
                        "Fecha_Salida": "Fecha Envío",
                        "Receptor": "Cliente / Receptor",
                        "SKU": "Código (SKU)",
                        "Nombre": "Descripción Comercial",
                        "Cantidad": "Piezas"
                    })
                    
                    columnas_mostrar = ["Remisión", "Fecha Envío", "Cliente / Receptor", "Tarima", "Código (SKU)", "Descripción Comercial", "Piezas", "PO", "Proyecto"]
                    columnas_mostrar = [c for c in columnas_mostrar if c in df_cruce.columns]
                    
                    df_mostrar = df_cruce[columnas_mostrar].copy()
                    
                    for col in df_mostrar.columns:
                        if col != "Piezas":
                            df_mostrar[col] = df_mostrar[col].astype(str)
                    df_mostrar["Piezas"] = pd.to_numeric(df_mostrar["Piezas"], errors='coerce').fillna(0).astype(int)
                    
                    total_pzs = df_mostrar["Piezas"].sum()
                    
                    st.write("---")
                    col_met1, col_met2 = st.columns(2)
                    col_met1.metric(f"📦 Total Piezas Remisionadas a {receptor_sel}", f"{total_pzs:,} Pzs")
                    col_met2.metric("🚚 Total Remisiones Relacionadas", df_mostrar["Remisión"].nunique())
                    
                    st.write("#### Detalle Operativo")
                    st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
                    
                    with col_rec2:
                        st.write("") # Alineación vertical
                        csv = df_mostrar.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="📥 Descargar CSV",
                            data=csv,
                            file_name=f"Reporte_Receptor_{receptor_sel}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                else:
                    st.warning("No hay registros de piezas para las tarimas enviadas a este receptor.")
            else:
                st.warning("No hay tarimas asociadas a las remisiones de este receptor.")
    else:
        st.info("No hay datos suficientes (Falta Base de Remisiones o Detalle de Tarimas).")

elif opcion_menu == "⚙️ Mantenimiento y Catálogos":
    st.title("⚙️ Panel de Mantenimiento Avanzado del Sistema")
    st.warning("⚠️ Acción Crítica: Las modificaciones realizadas impactan directamente en los archivos de GitHub.")
    
    # Inicialización del catálogo básico de personal operativo en memoria
    if "BD_Lideres" not in st.session_state:
        st.session_state.BD_Lideres = pd.DataFrame([
            {"ID_Lider": "LID-01", "Nombre_Lider": "Jesus Morales", "Area": "Metales", "Estatus": "Activo"}
        ])

    # REVISA QUE TU LÍNEA DE PESTAÑAS SUPERIOR TENGA ESTA ASIGNACIÓN:
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📝 Ajustar Cantidades", "👥 Catálogo de Líderes", "⚠️ Purga de Datos", "📦 Catálogo de Artículos", "🔢 Contador de Tarimas", "🖼️ Carpeta de Imágenes"])



    # --- SUB-MÓDULO 1: MODIFICACIÓN DIRECTA DE DETALLES DE INVENTARIO ---
    with tab1:
        st.subheader("✏️ Edición Rápida de Inventario (Detalle Tarimas)")
        if not st.session_state.BD_Detalle_Tarimas.empty:
            st.markdown("Haga doble clic sobre cualquier celda permitida (**SKU, PO, Parcialidad, Descripcion o Cantidad**) para corregir errores:")
            
            # Editor interactivo bloqueando únicamente llaves primarias y relacionales del sistema
            df_editable = st.data_editor(
                st.session_state.BD_Detalle_Tarimas, 
                use_container_width=True, 
                disabled=["ID_Detalle", "ID_Tarima", "Proyecto"], 
                hide_index=True, 
                key="editor_mantenimiento_cantidades_final"
            )
            
            if st.button("💾 Guardar Cambios de Inventario en GitHub"):
                st.session_state.BD_Detalle_Tarimas = df_editable
                subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                st.success("✅ ¡Inventario corregido y sincronizado con éxito!"); st.rerun()
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
        if st.session_state.rol != "Administrador":
            st.error("🔒 Área Bloqueada: Solo el Administrador del sistema tiene permisos para purgar o eliminar registros de las bases de datos.")
        else:
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
                        import os
                        if os.path.exists("consecutivo_override.txt"):
                            try:
                                os.remove("consecutivo_override.txt")
                            except Exception:
                                pass
                        st.session_state["siguiente_numero_tpm"] = 1
                        st.success("💥 Sistema purgado por completo a ceros de forma masiva."); st.rerun()
            else:
                st.markdown("### 📦 1. Eliminar Tarimas del Inventario")
                if not st.session_state.BD_Tarimas.empty:
                    df_tar_vista = st.session_state.BD_Tarimas.copy().drop(columns=["Es_Nueva"], errors="ignore")
                    sel_tarimas = st.dataframe(df_tar_vista, use_container_width=True, on_select="rerun", selection_mode="multi-row", key="tabla_purga_tarimas_final_f")
                    filas_tar = sel_tarimas.get("selection", {}).get("rows", [])
                    if filas_tar:
                        # Filtrar índices válidos para evitar errores de desfase de Streamlit en la recarga
                        filas_tar_validas = [i for i in filas_tar if i < len(df_tar_vista)]
                        ids_tar_eliminar = df_tar_vista.iloc[filas_tar_validas]['ID_Tarima'].tolist() if filas_tar_validas else []
                        
                        if st.button("🗑️ Eliminar Tarimas Seleccionadas") and ids_tar_eliminar:
                            st.session_state.BD_Tarimas = st.session_state.BD_Tarimas[~st.session_state.BD_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                            st.session_state.BD_Detalle_Tarimas = st.session_state.BD_Detalle_Tarimas[~st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(ids_tar_eliminar)]
                            subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                            subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                            st.session_state["siguiente_numero_tpm"] = obtener_siguiente_consecutivo_tpm()
                            st.success("✅ Tarimas removidas."); st.rerun()
                else: st.write("No hay tarimas registradas.")
                
                st.write("---")
                st.markdown("### 🚚 2. Eliminar Remisiones de Salida")
                if not st.session_state.BD_Datos_Generales_Remision.empty:
                    df_rem_purga = st.session_state.BD_Datos_Generales_Remision.copy()
                    if 'Tarimas_Asociadas' in df_rem_purga.columns:
                        df_rem_purga['Tarimas_Asociadas'] = df_rem_purga['Tarimas_Asociadas'].astype(str)
                    sel_remisiones = st.dataframe(df_rem_purga, use_container_width=True, on_select="rerun", selection_mode="multi-row", key="tabla_purga_remisiones_final_f")
                    filas_rem = sel_remisiones.get("selection", {}).get("rows", [])
                    if filas_rem:
                        # Filtrar índices válidos para evitar errores de desfase de Streamlit en la recarga
                        filas_rem_validas = [i for i in filas_rem if i < len(st.session_state.BD_Datos_Generales_Remision)]
                        if filas_rem_validas:
                            ids_rem_eliminar = st.session_state.BD_Datos_Generales_Remision.iloc[filas_rem_validas]['Folio_Remision'].tolist()
                            tarimas_afectadas = []
                            import ast
                            for idx in filas_rem_validas:
                                raw_tarimas = st.session_state.BD_Datos_Generales_Remision.iloc[idx]['Tarimas_Asociadas']
                                if isinstance(raw_tarimas, str):
                                    try:
                                        t_list = ast.literal_eval(raw_tarimas)
                                    except Exception:
                                        t_list = [raw_tarimas]
                                elif isinstance(raw_tarimas, list):
                                    t_list = raw_tarimas
                                else:
                                    t_list = []
                                for t in t_list:
                                    tarimas_afectadas.append(str(t).strip())
                        else:
                            ids_rem_eliminar = []
                            tarimas_afectadas = []
    
                        if st.button("🗑️ Eliminar Remisiones Seleccionadas") and ids_rem_eliminar:
                            st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(tarimas_afectadas), 'Estatus'] = 'Disponible'
                            st.session_state.BD_Datos_Generales_Remision = st.session_state.BD_Datos_Generales_Remision[~st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].isin(ids_rem_eliminar)]
                            subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                            subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                            st.success("✅ Remisiones eliminadas y bultos reactivados."); st.rerun()
                else: st.write("No hay remisiones registradas.")
                
                st.write("---")
                st.markdown("### 🔄 3. Sincronización y Reparación de Estatus de Tarimas")
                st.info("Utilice esta herramienta para buscar tarimas marcadas como 'Remesadas' pero que no pertenecen a ninguna remisión activa (por ejemplo, remisiones borradas manualmente), y regresarlas al estatus de 'Disponible'.")
                if st.button("🔄 Sincronizar y Reparar Estatus de Tarimas"):
                    corregidas = sincronizar_estatus_tarimas(auto_save=True)
                    if corregidas > 0:
                        st.success(f"✅ ¡Operación completada! Se corrigieron y liberaron {corregidas} tarimas huérfanas.")
                        st.rerun()
                    else:
                        st.info("ℹ️ No se detectaron tarimas con estatus incorrecto. Todo está sincronizado correctamente.")
   
    # --- SUB-MÓDULO 4: ADMINISTRACIÓN MASIVA DEL CATÁLOGO DE ARTÍCULOS ---
# --- SUB-MÓDULO 4: ADMINISTRACIÓN MASIVA DEL CATÁLOGO DE ARTÍCULOS ---
# --- SUB-MÓDULO 4: ADMINISTRACIÓN MASIVA DEL CATÁLOGO DE ARTÍCULOS ---
# --- SUB-MÓDULO 4: ADMINISTRACIÓN MASIVA Y EDICIÓN DEL CATÁLOGO DE ARTÍCULOS ---
    with tab4:
        st.subheader("📦 Administración y Sincronización del Catálogo de Artículos")
        st.markdown("Utilice este panel para actualizar de forma masiva mediante archivos o editar directamente registros específicos en caliente:")
    
        # Forzamos la presencia local de librerías críticas
        import io
        import pandas as pd
        from openpyxl.worksheet.datavalidation import DataValidation
    
        # =============================================================================
        # BLOQUE A: CARGA MASIVA Y PLANTILLAS
        # =============================================================================
        with st.expander("📥 Carga Masiva mediante Excel (Subir / Descargar Plantilla)", expanded=False):
            c_art1, c_art2 = st.columns(2)
            
            with c_art1:
                st.write("##### Obtener Plantilla Estructurada")
                columnas_oficiales = ["SKU", "Nombre", "Calibre_Espesor", "Dimensiones_Pieza", "Acabado_Superficial"]
                df_plantilla_art = pd.DataFrame(columns=columnas_oficiales)
    
                buf_p_art = io.BytesIO()
                with pd.ExcelWriter(buf_p_art, engine='openpyxl') as p_writer:
                    df_plantilla_art.to_excel(p_writer, index=False, sheet_name='Datos_Sistema')
                    worksheet = p_writer.sheets['Datos_Sistema']
                    worksheet.column_dimensions['A'].width = 20
                    worksheet.column_dimensions['B'].width = 25
                    worksheet.column_dimensions['C'].width = 20
                    worksheet.column_dimensions['D'].width = 25
                    worksheet.column_dimensions['E'].width = 20
    
                    opciones_calibres = '"10GA,12GA,14GA,16GA,10GACR,12GACR,14GACR,16GACR","125AL","250AL","188AL"'
                    dv_calibre = DataValidation(type="list", formula1=opciones_calibres, allow_blank=True)
                    worksheet.add_data_validation(dv_calibre)
                    dv_calibre.add("C2:C2000")
    
                    opciones_acabados = '"Decapado,Ansi 61,Galvanizado,Otro"'
                    dv_acabado = DataValidation(type="list", formula1=opciones_acabados, allow_blank=True)
                    worksheet.add_data_validation(dv_acabado)
                    dv_acabado.add("E2:E2000")
    
                buf_p_art.seek(0)
                st.download_button(
                    label="📊 Descargar Plantilla Base Artículos (.xlsx)",
                    data=buf_p_art.getvalue(),
                    file_name="Plantilla_Maestra_Articulos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_download_plantilla_maestra_articulos_sgc_v2"
                )

                st.write("---")
                st.write("##### Artículos en Tarimas Sin Registro en Catálogo")

                skus_en_tarimas = set()
                if "BD_Detalle_Tarimas" in st.session_state and not st.session_state.BD_Detalle_Tarimas.empty:
                    skus_en_tarimas = set(st.session_state.BD_Detalle_Tarimas['SKU'].astype(str).str.strip().dropna().unique())

                skus_en_articulos = set()
                if "BD_Articulos" in st.session_state and not st.session_state.BD_Articulos.empty:
                    skus_en_articulos = set(st.session_state.BD_Articulos['SKU'].astype(str).str.strip().dropna().unique())

                skus_sin_registro = sorted(list(skus_en_tarimas - skus_en_articulos))
                # Excluir valores vacíos o de selección
                skus_sin_registro = [s for s in skus_sin_registro if s and s.upper() not in ["TODOS", "SELECCIONE UN SKU..."]]

                if len(skus_sin_registro) > 0:
                    st.info(f"🔍 Se encontraron **{len(skus_sin_registro)}** artículos en tarimas que no están registrados en el Catálogo.")
                else:
                    st.success("🎉 ¡Todos los artículos en tarimas están registrados en el Catálogo!")

                df_sin_registro = pd.DataFrame(columns=columnas_oficiales)
                df_sin_registro["SKU"] = skus_sin_registro

                buf_sin_reg = io.BytesIO()
                with pd.ExcelWriter(buf_sin_reg, engine='openpyxl') as p_writer:
                    df_sin_registro.to_excel(p_writer, index=False, sheet_name='Datos_Sistema')
                    worksheet = p_writer.sheets['Datos_Sistema']
                    worksheet.column_dimensions['A'].width = 20
                    worksheet.column_dimensions['B'].width = 25
                    worksheet.column_dimensions['C'].width = 20
                    worksheet.column_dimensions['D'].width = 25
                    worksheet.column_dimensions['E'].width = 20

                    opciones_calibres = '"10GA,12GA,14GA,16GA,10GACR,12GACR,14GACR,16GACR","125AL","250AL","188AL"'
                    dv_calibre = DataValidation(type="list", formula1=opciones_calibres, allow_blank=True)
                    worksheet.add_data_validation(dv_calibre)
                    dv_calibre.add("C2:C2000")

                    opciones_acabados = '"Decapado,Ansi 61,Galvanizado,Otro"'
                    dv_acabado = DataValidation(type="list", formula1=opciones_acabados, allow_blank=True)
                    worksheet.add_data_validation(dv_acabado)
                    dv_acabado.add("E2:E2000")

                buf_sin_reg.seek(0)

                st.download_button(
                    label="📥 Articulos Sin Registro",
                    data=buf_sin_reg.getvalue(),
                    file_name="Plantilla_Articulos_Sin_Registro.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_download_articulos_sin_registro",
                    disabled=len(skus_sin_registro) == 0,
                    use_container_width=True
                )
    
            with c_art2:
                st.write("##### Cargar Catálogo Completo")
                arch_articulos = st.file_uploader("Suba la Plantilla de Artículos Rellenada:", type=["xlsx"], key="uploader_articulos_masivo_f")
        
                if arch_articulos:
                    if st.button("🔄 Procesar e Integrar Catálogo en GitHub", use_container_width=True):
                        try:
                            df_art_excel = pd.read_excel(arch_articulos)
                            columnas_requeridas = ["SKU", "Nombre", "Calibre_Espesor", "Dimensiones_Pieza", "Acabado_Superficial"]
        
                            # Verificación manual y limpia de columnas para evitar fallas de sintaxis
                            columnas_correctas = True
                            for col in columnas_requeridas:
                                if col not in df_art_excel.columns:
                                    columnas_correctas = False
        
                            if not columnas_correctas:
                                st.error("❌ Error: Columnas incompatibles. Use la estructura oficial.")
                            else:
                                # Limpieza del archivo subido
                                df_art_excel = df_art_excel.dropna(subset=["SKU"])
                                df_art_excel["SKU"] = df_art_excel["SKU"].astype(str).str.strip().str.upper()
        
                                # Integración inteligente: Conserva anteriores y suma/actualiza los nuevos
                                if "BD_Articulos" in st.session_state and not st.session_state.BD_Articulos.empty:
                                    df_anterior = st.session_state.BD_Articulos.copy()
                                    df_anterior["SKU"] = df_anterior["SKU"].astype(str).str.strip().str.upper()
                                    
                                    # Eliminamos del catálogo anterior los SKUs que vienen en el archivo nuevo para actualizarlos
                                    df_anterior = df_anterior[~df_anterior["SKU"].isin(df_art_excel["SKU"])]
                                    
                                    # Unimos los registros viejos con los nuevos
                                    st.session_state.BD_Articulos = pd.concat([df_anterior, df_art_excel], ignore_index=True)
                                else:
                                    st.session_state.BD_Articulos = df_art_excel
        
                                exito_subida = subir_excel_a_github("BD_Articulos.xlsx", st.session_state.BD_Articulos)
        
                                if exito_subida:
                                    st.success("✅ ¡Catálogo actualizado e integrado en GitHub exitosamente!")
                                    st.rerun()
                                else:
                                    st.error("❌ Error al subir el archivo al repositorio de GitHub.")
                        except Exception as e:
                            st.error(f"⚠️ Error crítico: {e}")
        
            st.write("---")
        
            # =============================================================================
            # BLOQUE B: EDITOR EN CALIENTE TIPO EXCEL CON ELIMINACIÓN HABILITADA
            # =============================================================================
            st.write("##### ✏️ Edición Directa, Alta y Baja del Catálogo Maestro")
            st.info("💡 **Guía de uso:**\n"
                    "* **Editar:** Haga doble clic sobre cualquier celda.\n"
                    "* **Eliminar:** Seleccione la fila desde el extremo izquierdo y presione la tecla **Supr/Delete** en su teclado.\n"
                    "* **Nota:** El SKU actúa como llave relacional y está bloqueado para edición, pero sí puede eliminar la fila completa si ya no se usa.")
        
            if "BD_Articulos" in st.session_state and not st.session_state.BD_Articulos.empty:
                df_art_editable = st.data_editor(
                    st.session_state.BD_Articulos,
                    use_container_width=True,
                    disabled=["SKU"],
                    hide_index=True,
                    num_rows="dynamic",  # Permite añadir (+) y eliminar filas nativas
                    key="editor_mantenimiento_articulos_directo_v3",
                    column_config={
                        "SKU": st.column_config.TextColumn("SKU (Bloqueado/Llave)"),
                        "Nombre": st.column_config.TextColumn("Descripción Comercial"),
                        "Calibre_Espesor": st.column_config.SelectboxColumn("Calibre / Espesor", options=["10GA", "12GA", "14GA", "16GA", "10GACR", "12GACR", "14GACR", "16GACR", "125AL", "250AL", "188AL"]),
                        "Dimensiones_Pieza": st.column_config.TextColumn("Dimensiones"),
                        "Acabado_Superficial": st.column_config.SelectboxColumn("Acabado Superficial", options=["Decapado", "Ansi 61", "Galvanizado", "Otro"])
                    }
                )
        
                # Botón de guardado y persistencia en GitHub
                if st.button("💾 Guardar Cambios y Aplicar Bajas en GitHub", use_container_width=True):
                    # Limpieza preventiva de datos antes de subir
                    df_art_final = df_art_editable.dropna(subset=["SKU"])
                    df_art_final["SKU"] = df_art_final["SKU"].astype(str).str.strip().str.upper()
        
                    # Actualizar el estado de la sesión local
                    st.session_state.BD_Articulos = df_art_final
        
                    # Sincronizar y sobreescribir en la nube
                    exito_guardado = subir_excel_a_github("BD_Articulos.xlsx", st.session_state.BD_Articulos)
        
                    if exito_guardado:
                        st.success("💥 ¡Catálogo de artículos actualizado! Las modificaciones y bajas se sincronizaron con éxito.")
                        st.rerun()
                    else:
                        st.error("❌ Error de comunicación: No se pudieron plasmar los cambios en el repositorio.")
            else:
                st.warning("⚠️ No existen registros activos en el catálogo de artículos para desplegar el editor.")
    # =============================================================================
    # PESTAÑA 5: CONFIGURADOR DEL FOLIO CONSECUTIVO TPM
    # =============================================================================
    try:
        # Intentamos usar la variable de la quinta pestaña si existe
        with tab5:
            st.write("##### 🔢 Inicializar o Cambiar Consecutivo de Tarimas (TPM)")
            st.info("💡 **Guía de uso:** Defina el número numérico en el que desea que continúe la siguiente tarima que genere el sistema (Ejemplo: Si pone 56, la siguiente será TPM-0056).")
    
            if "siguiente_numero_tpm" not in st.session_state or st.session_state["siguiente_numero_tpm"] is None:
                st.session_state["siguiente_numero_tpm"] = obtener_siguiente_consecutivo_tpm()
    
            nuevo_consecutivo = st.number_input(
                "Indique el siguiente número de folio a generar:",
                min_value=1, max_value=9999,
                value=int(st.session_state["siguiente_numero_tpm"]),
                step=1, key="input_consecutivo_manual_tpm"
            )
    
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("💾 Guardar Nuevo Consecutivo de Folio", use_container_width=True):
                    st.session_state["siguiente_numero_tpm"] = nuevo_consecutivo
                    try:
                        with open("consecutivo_override.txt", "w", encoding="utf-8") as f:
                            f.write(str(nuevo_consecutivo))
                    except Exception as e:
                        st.warning(f"⚠️ No se pudo guardar el consecutivo en disco: {e}")
                    st.success(f"💥 ¡Consecutivo actualizado! La próxima tarima nueva se creará con el folio: **TPM-{nuevo_consecutivo:04d}**")
                    st.rerun()
            with col_btn2:
                if st.button("🔄 Restablecer a Consecutivo Automático", use_container_width=True):
                    import os
                    if os.path.exists("consecutivo_override.txt"):
                        try:
                            os.remove("consecutivo_override.txt")
                        except Exception:
                            pass
                    st.session_state["siguiente_numero_tpm"] = obtener_siguiente_consecutivo_tpm()
                    st.success("🔄 Restablecido al siguiente consecutivo automático de la base de datos.")
                    st.rerun()
        with tab6:
            renderizar_explorador_imagenes()
    except NameError:
        # Si tab5 no existe, desplegamos un contenedor nativo expandible para que no rompa la aplicación
        with st.expander("🔢 Contador y Consecutivo de Tarimas (TPM)", expanded=True):
            st.write("##### 🔢 Inicializar o Cambiar Consecutivo de Tarimas (TPM)")
            st.info("💡 **Guía de uso:** Defina el número numérico en el que desea que continúe la siguiente tarima que genere el sistema (Ejemplo: Si pone 56, la siguiente será TPM-0056).")
    
            if "siguiente_numero_tpm" not in st.session_state or st.session_state["siguiente_numero_tpm"] is None:
                st.session_state["siguiente_numero_tpm"] = obtener_siguiente_consecutivo_tpm()
    
            nuevo_consecutivo = st.number_input(
                "Indique el siguiente número de folio a generar:",
                min_value=1, max_value=9999,
                value=int(st.session_state["siguiente_numero_tpm"]),
                step=1, key="input_consecutivo_manual_tpm_fallback"
            )
    
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("💾 Guardar Nuevo Consecutivo de Folio ", use_container_width=True):
                    st.session_state["siguiente_numero_tpm"] = nuevo_consecutivo
                    try:
                        with open("consecutivo_override.txt", "w", encoding="utf-8") as f:
                            f.write(str(nuevo_consecutivo))
                    except Exception as e:
                        st.warning(f"⚠️ No se pudo guardar el consecutivo en disco: {e}")
                    st.success(f"💥 ¡Consecutivo actualizado! La próxima tarima nueva se creará con el folio: **TPM-{nuevo_consecutivo:04d}**")
                    st.rerun()
            with col_btn2:
                if st.button("🔄 Restablecer a Consecutivo Automático ", use_container_width=True):
                    import os
                    if os.path.exists("consecutivo_override.txt"):
                        try:
                            os.remove("consecutivo_override.txt")
                        except Exception:
                            pass
                    st.session_state["siguiente_numero_tpm"] = obtener_siguiente_consecutivo_tpm()
                    st.success("🔄 Restablecido al siguiente consecutivo automático de la base de datos.")
                    st.rerun()

        with st.expander("🖼️ Carpeta de Imágenes de Artículos (Mantenimiento)", expanded=False):
            renderizar_explorador_imagenes()

# =============================================================================
# 15. CARGA MASIVA HISTÓRICA
# =============================================================================
elif opcion_menu == "🕰️ Carga Histórica":
    st.title("🕰️ Carga Masiva de Remisiones y Tarimas Históricas")
    st.markdown("Esta sección permite cargar un archivo de Excel para dar de alta tarimas antiguas y registrarlas inmediatamente como remisiones enviadas, consolidando tu inventario histórico en un solo paso.")
    
    st.warning("⚠️ **ATENCIÓN:** El archivo de Excel debe contener estrictamente las siguientes columnas: **Tarima, Producto/SKU, PO, Proyecto, Parcialidad, Descripcion, Cantidad, Fecha de Remisión, Folio de Remisión**.")
    
    with st.form("form_carga_historica"):
        col1, col2 = st.columns(2)
        with col1:
            st.info("💡 **NUEVO:** El Folio de Remisión (Ej. E0001) ahora se leerá automáticamente de tu archivo Excel para cada tarima, permitiéndote subir múltiples remisiones en un solo documento.")
            receptor_input = st.text_input("Receptor (Destino para todas):", value="Galvatec Industrias")
        with col2:
            direccion_input = st.text_input("Dirección del Receptor:", value="Prol. Valle Guadiana 919, Parque Industrial II, 35078 Gómez Palacio, Dgo.")
            
        archivo_cargado = st.file_uploader("📥 Selecciona el archivo Excel con el historial", type=['xlsx'])
        btn_procesar = st.form_submit_button("🚀 Procesar Carga Histórica")
        
    if btn_procesar:
        if not archivo_cargado:
            st.error("❌ Debes subir el archivo Excel.")
        else:
            try:
                df_hist = pd.read_excel(archivo_cargado)
                # Soportar ambas formas del encabezado: "Folio de Remisión" o "Folio Remisión"
                if "Folio Remisión" in df_hist.columns and "Folio de Remisión" not in df_hist.columns:
                    df_hist = df_hist.rename(columns={"Folio Remisión": "Folio de Remisión"})
                    
                cols_requeridas = ["Tarima", "Producto/SKU", "PO", "Proyecto", "Parcialidad", "Descripcion", "Cantidad", "Fecha de Remisión", "Folio de Remisión"]
                
                # Validar columnas
                faltantes = [c for c in cols_requeridas if c not in df_hist.columns]
                if faltantes:
                    st.error(f"❌ El archivo no tiene el formato correcto. Faltan las columnas: {', '.join(faltantes)}")
                else:
                    # Auditoría de Datos
                    st.session_state['df_historico_auditado'] = df_hist
                    st.session_state['receptor_auditado'] = receptor_input
                    st.session_state['direccion_auditada'] = direccion_input
            except Exception as e:
                st.error(f"❌ Ocurrió un error al leer el archivo: {e}")

    # Sección de Auditoría y Confirmación
    if 'df_historico_auditado' in st.session_state:
        df_hist = st.session_state['df_historico_auditado']
        receptor_input = st.session_state['receptor_auditado']
        direccion_input = st.session_state['direccion_auditada']
        
        st.markdown("---")
        st.subheader("📋 Auditoría de Información a Cargar")
        
        tarimas_excel = df_hist['Tarima'].dropna().unique().tolist()
        folios_excel = df_hist['Folio de Remisión'].dropna().unique().tolist()
        skus_excel = df_hist['Producto/SKU'].dropna().unique().tolist()
        
        # Validar si los folios ya existen
        folios_existentes = []
        if not st.session_state.BD_Datos_Generales_Remision.empty:
            folios_bd = st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].astype(str).str.strip().tolist()
            for f in folios_excel:
                if str(f).strip() in folios_bd:
                    folios_existentes.append(str(f).strip())
                    
        # Validar si los SKUs existen
        skus_no_encontrados = []
        if "BD_Articulos" in st.session_state and not st.session_state.BD_Articulos.empty:
            skus_bd = st.session_state.BD_Articulos['SKU'].astype(str).str.strip().tolist()
            for s in skus_excel:
                if str(s).strip() not in skus_bd:
                    skus_no_encontrados.append(str(s).strip())
                    
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.metric("Total Tarimas a Crear", len(tarimas_excel))
            st.metric("Total Remisiones a Generar", len(folios_excel))
        with col_res2:
            st.metric("Total de Partidas (Filas)", len(df_hist))
            st.metric("SKUs únicos procesados", len(skus_excel))
            
        hay_conflictos = False
        if folios_existentes:
            hay_conflictos = True
            st.error(f"⚠️ **ALERTA DE CONFLICTO:** Los siguientes folios ya existen en el sistema: {', '.join(folios_existentes)}. Por seguridad, no puedes cargar duplicados.")
            
        if skus_no_encontrados:
            st.warning(f"⚠️ **ADVERTENCIA:** Se encontraron {len(skus_no_encontrados)} SKUs que no están en el Catálogo de Artículos. (Se pueden cargar, pero la descripción no se enriquecerá automáticamente).")
            
        if hay_conflictos:
            st.error("❌ Debes corregir el archivo Excel y volver a subirlo para evitar duplicar folios.")
            if st.button("Cancelar Carga"):
                del st.session_state['df_historico_auditado']
                st.rerun()
        else:
            st.success("✅ **Auditoría superada sin conflictos graves.** La información está lista para consolidarse.")
            
            if st.button("✅ Confirmar y Ejecutar Carga Masiva", use_container_width=True):
                with st.spinner("Procesando la carga en las bases de datos..."):
                    try:
                        mapa_tpm = {}
                        nuevo_id = obtener_siguiente_consecutivo_tpm()
                        for t in tarimas_excel:
                            mapa_tpm[t] = f"TPM-{nuevo_id:04d}"
                            nuevo_id += 1
                            
                        st.session_state["siguiente_numero_tpm"] = nuevo_id
                        
                        # 3. Preparar BD_Tarimas
                        nuevas_tarimas = []
                        for t in tarimas_excel:
                            sub_df = df_hist[df_hist['Tarima'] == t]
                            fecha_str = str(sub_df.iloc[0]['Fecha de Remisión']).split(' ')[0]
                            nuevas_tarimas.append({
                                "ID_Tarima": mapa_tpm[t],
                                "Tarima_Origen_Excel": t,
                                "Fecha_Creacion": fecha_str,
                                "Ubicacion_Actual": "Planta Rio",
                                "Creado_Por": "Carga Historica",
                                "Tipo_Tarima": "Mixta",
                                "Estatus": "Remesada",
                                "Es_Nueva": "No"
                            })
                        df_nuevas_tarimas = pd.DataFrame(nuevas_tarimas)
                        st.session_state.BD_Tarimas = pd.concat([st.session_state.BD_Tarimas, df_nuevas_tarimas], ignore_index=True)
                        
                        # 4. Preparar BD_Detalle_Tarimas
                        nuevos_detalles = []
                        id_det = 1
                        if not st.session_state.BD_Detalle_Tarimas.empty:
                            try:
                                id_det = int(st.session_state.BD_Detalle_Tarimas['ID_Detalle'].max()) + 1
                            except:
                                id_det = len(st.session_state.BD_Detalle_Tarimas) + 1
                                
                        for _, row in df_hist.iterrows():
                            t_excel = row['Tarima']
                            nuevos_detalles.append({
                                "ID_Detalle": id_det,
                                "ID_Tarima": mapa_tpm[t_excel],
                                "SKU": row['Producto/SKU'],
                                "PO": row['PO'],
                                "Proyecto": row['Proyecto'],
                                "Parcialidad": row['Parcialidad'],
                                "Descripcion": row['Descripcion'],
                                "Cantidad": row['Cantidad']
                            })
                            id_det += 1
                        df_nuevos_detalles = pd.DataFrame(nuevos_detalles)
                        st.session_state.BD_Detalle_Tarimas = pd.concat([st.session_state.BD_Detalle_Tarimas, df_nuevos_detalles], ignore_index=True)
                        
                        # 5. Preparar BD_Datos_Generales_Remision (Agrupando por Folio)
                        import json
                        nuevas_remisiones = []
                        
                        id_rem = 1
                        if not st.session_state.BD_Datos_Generales_Remision.empty:
                            try:
                                id_rem = int(st.session_state.BD_Datos_Generales_Remision['ID_Remision'].max()) + 1
                            except:
                                id_rem = len(st.session_state.BD_Datos_Generales_Remision) + 1
                                
                        for folio in folios_excel:
                            sub_df_folio = df_hist[df_hist['Folio de Remisión'] == folio]
                            fecha_rem = str(sub_df_folio.iloc[0]['Fecha de Remisión']).split(' ')[0]
                            tarimas_de_este_folio_excel = sub_df_folio['Tarima'].dropna().unique().tolist()
                            tarimas_asignadas_list = [mapa_tpm[t] for t in tarimas_de_este_folio_excel]
                            
                            nuevas_remisiones.append({
                                "ID_Remision": id_rem,
                                "Folio_Remision": str(folio).strip(),
                                "Fecha_Hora_Salida": fecha_rem,
                                "Nombre_Emisor": "Carga Historica",
                                "Direccion_Emisor": "N/A",
                                "Nombre_Receptor": receptor_input.strip(),
                                "Direccion_Receptor": direccion_input.strip(),
                                "Tarimas_Asociadas": json.dumps(tarimas_asignadas_list)
                            })
                            id_rem += 1
                            
                        df_nuevas_remisiones = pd.DataFrame(nuevas_remisiones)
                        st.session_state.BD_Datos_Generales_Remision = pd.concat([st.session_state.BD_Datos_Generales_Remision, df_nuevas_remisiones], ignore_index=True)
                        
                        # 6. Guardar todo
                        subir_excel_a_github("BD_Tarimas.xlsx", st.session_state.BD_Tarimas)
                        subir_excel_a_github("BD_Detalle_Tarimas.xlsx", st.session_state.BD_Detalle_Tarimas)
                        subir_excel_a_github("BD_Datos_Generales_Remision.xlsx", st.session_state.BD_Datos_Generales_Remision)
                        
                        try:
                            with open("consecutivo_override.txt", "w", encoding="utf-8") as f:
                                f.write(str(nuevo_id))
                        except:
                            pass
                            
                        del st.session_state['df_historico_auditado']
                        st.success(f"✅ ¡Carga Exitosa! Se crearon {len(tarimas_excel)} tarimas en {len(folios_excel)} remisiones.")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Ocurrió un error crítico durante la carga: {e}")
    
