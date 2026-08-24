import streamlit as st
import os
import pandas as pd
import datetime
import io
import requests
from PIL import Image
import glob

# 1. CONFIGURACIÓN E INTERFAZ BASE RESPONSIVA (Portal de Consulta Interna)
st.set_page_config(
    page_title="Consulta de Inventario y SKUs — SIGRAMA Metales",
    layout="wide",
    page_icon="🔍"
)

# REPOSITORIO DE DATOS
REPO_OWNER = "jesusalbertomoraleslopez-byte"
REPO_NAME = "remisiones-de-materiales"
BRANCH = "main"

# Helper para secrets
def obtener_secret(key, default=None):
    try:
        if hasattr(st, 'secrets') and st.secrets is not None:
            if key in st.secrets:
                val = st.secrets[key]
                return val if val is not None else default
    except Exception:
        pass
    return default

# Estilos CSS Corporativos (Industria SIGRAMA)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;700&family=Questrial&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Questrial', sans-serif !important;
        background-color: #FFFFFF !important;
    }

    h1, h2, h3, h4, h5, h6 {
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        color: #111111 !important;
    }

    .main-header {
        background: linear-gradient(135deg, #111111 0%, #1E293B 100%);
        padding: 20px 25px;
        border-radius: 10px;
        border-left: 6px solid #EC2024;
        margin-bottom: 20px;
        color: #FFFFFF;
    }

    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
    }

    .status-disponible {
        background-color: #DCFCE7;
        color: #15803D;
        font-weight: bold;
        padding: 4px 8px;
        border-radius: 4px;
    }

    .status-remesada {
        background-color: #DBEAFE;
        color: #1D4ED8;
        font-weight: bold;
        padding: 4px 8px;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# 2. CARGA DE DATOS DESDE GITHUB CON CACHÉ DE MEMORIA
@st.cache_data(ttl=120, max_entries=10)
def cargar_excel_desde_github(file_name):
    """Carga archivos Excel desde la API de GitHub o fallback a disco local."""
    try:
        url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{file_name}?ref={BRANCH}&ts={int(__import__('time').time())}"
        headers = {"Cache-Control": "no-cache", "Accept": "application/vnd.github.v3.raw"}
        token = obtener_secret("github_token")
        if token:
            headers["Authorization"] = f"token {token}"
            
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            archivo_bytes = res.content
            try:
                return pd.read_excel(io.BytesIO(archivo_bytes), sheet_name='Datos_Sistema')
            except Exception:
                return pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=0)
    except Exception:
        pass

    if os.path.exists(file_name):
        try:
            try:
                return pd.read_excel(file_name, sheet_name='Datos_Sistema')
            except Exception:
                return pd.read_excel(file_name, sheet_name=0)
        except Exception:
            pass
    return pd.DataFrame()

# Cargar Bases de Datos
df_articulos = cargar_excel_desde_github("BD_Articulos.xlsx")
df_skus_aut = cargar_excel_desde_github("BD_SKUs_Autorizados.xlsx")
df_tarimas = cargar_excel_desde_github("BD_Tarimas.xlsx")
df_detalle = cargar_excel_desde_github("BD_Detalle_Tarimas.xlsx")
df_remisiones = cargar_excel_desde_github("BD_Datos_Generales_Remision.xlsx")

# --- ENCABEZADO ---
col_logo, col_title = st.columns([1, 4])
with col_logo:
    if os.path.exists("logo_sigrama.png"):
        st.image("logo_sigrama.png", use_container_width=True)
    else:
        st.title("🏭")
with col_title:
    st.markdown("""
    <div class="main-header">
        <h2 style="margin:0; color:#FFFFFF; font-size:24px;">🔍 PORTAL DE CONSULTA DE INVENTARIO Y UBIACIONES</h2>
        <p style="margin:4px 0 0 0; color:#94A3B8; font-size:14px;">
            Industria SIGRAMA S.A. de C.V. — Verificación en tiempo real de disponibilidad, fotografías y ubicación física por SKU.
        </p>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# --- OBTENER LISTA DE SKUS DISPONIBLES ---
skus_set = set()
if not df_articulos.empty and 'SKU' in df_articulos.columns:
    skus_set.update(df_articulos['SKU'].dropna().astype(str).str.strip().str.upper())
if not df_skus_aut.empty and 'SKU' in df_skus_aut.columns:
    skus_set.update(df_skus_aut['SKU'].dropna().astype(str).str.strip().str.upper())
if not df_detalle.empty and 'SKU' in df_detalle.columns:
    skus_set.update(df_detalle['SKU'].dropna().astype(str).str.strip().str.upper())

lista_skus = sorted([s for s in skus_set if s not in ['', 'NAN', 'NONE']])

# --- BÚSQUEDA Y FILTROS ---
st.markdown("### 🔎 Buscar Material")
tab_sku, tab_proj, tab_po = st.tabs(["🏷️ Búsqueda por SKU", "📌 Búsqueda por Proyecto", "📄 Búsqueda por PO"])

sku_seleccionado = None
proyecto_seleccionado = None
po_seleccionada = None

with tab_sku:
    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        sku_input = st.selectbox(
            "Seleccione o escriba el Código de SKU:",
            options=["(Seleccione un SKU...)"] + lista_skus,
            key="lookup_sku_select"
        )
        if sku_input != "(Seleccione un SKU...)":
            sku_seleccionado = sku_input

    with col_s2:
        sku_manual = st.text_input("O escriba código manualmente:", placeholder="Ej: SKU-1234", key="lookup_sku_text")
        if sku_manual.strip():
            sku_seleccionado = sku_manual.strip().upper()

with tab_proj:
    if not df_detalle.empty and 'Proyecto' in df_detalle.columns:
        proyectos_disponibles = sorted([str(p).strip() for p in df_detalle['Proyecto'].dropna().unique() if str(p).strip() not in ['', 'nan', 'None']])
        proyecto_seleccionado = st.selectbox("Seleccione el Proyecto:", options=["(Seleccione un Proyecto...)"] + proyectos_disponibles, key="lookup_proj_select")
        if proyecto_seleccionado == "(Seleccione un Proyecto...)":
            proyecto_seleccionado = None

with tab_po:
    if not df_detalle.empty and 'PO' in df_detalle.columns:
        pos_disponibles = sorted([str(p).strip() for p in df_detalle['PO'].dropna().unique() if str(p).strip() not in ['', 'nan', 'None']])
        po_seleccionada = st.selectbox("Seleccione la Orden de Compra (PO):", options=["(Seleccione una PO...)"] + pos_disponibles, key="lookup_po_select")
        if po_seleccionada == "(Seleccione una PO...)":
            po_seleccionada = None

st.write("---")

# --- RESULTADOS DE CONSULTA ---
if sku_seleccionado:
    st.markdown(f"## 📦 Resultado para el SKU: `{sku_seleccionado}`")
    
    col_info, col_img = st.columns([2, 1])
    
    # 1. FICHA TÉCNICA DEL ARTÍCULO
    with col_info:
        st.markdown("#### 📋 Ficha Técnica del Artículo")
        
        nombre_art = "No especificado"
        calibre_art = "N/A"
        dims_art = "N/A"
        acabado_art = "N/A"
        
        if not df_articulos.empty and 'SKU' in df_articulos.columns:
            df_match = df_articulos[df_articulos['SKU'].astype(str).str.strip().str.upper() == sku_seleccionado]
            if not df_match.empty:
                art_row = df_match.iloc[0]
                nombre_art = str(art_row.get('Nombre', 'No especificado'))
                calibre_art = str(art_row.get('Calibre_Espesor', 'N/A'))
                dims_art = str(art_row.get('Dimensiones_Pieza', 'N/A'))
                acabado_art = str(art_row.get('Acabado_Superficial', 'N/A'))
                
        st.markdown(f"""
        - **Descripción Comercial:** `{nombre_art}`
        - **Calibre / Espesor:** `{calibre_art}`
        - **Dimensiones:** `{dims_art}`
        - **Material / Acabado:** `{acabado_art}`
        """)

    # 2. FOTOGRAFÍA DEL ARTÍCULO
    with col_img:
        st.markdown("#### 🖼️ Fotografía")
        matching_imgs = glob.glob(f"imagenes_articulos/{sku_seleccionado}(*.*")
        if matching_imgs and os.path.exists(matching_imgs[0]):
            try:
                img_pil = Image.open(matching_imgs[0])
                st.image(img_pil, caption=f"Imagen oficial de {sku_seleccionado}", use_container_width=True)
            except Exception:
                st.warning("⚠️ No se pudo renderizar la imagen cargada.")
        else:
            st.info("📷 *Sin imagen fotográfica asignada en catálogo.*")

    st.write("---")

    # 3. DETALLE DE INVENTARIO Y UBICACIÓN FÍSICA
    if not df_detalle.empty:
        df_sub_det = df_detalle[df_detalle['SKU'].astype(str).str.strip().str.upper() == sku_seleccionado].copy()
        
        if df_sub_det.empty:
            st.warning(f"⚠️ El SKU **{sku_seleccionado}** está registrado en el catálogo pero **NO cuenta con tarimas físicamente registradas** en inventario.")
        else:
            # Cruzar con BD_Tarimas para obtener Ubicación y Estatus
            if not df_tarimas.empty:
                df_sub_det = pd.merge(df_sub_det, df_tarimas[['ID_Tarima', 'Ubicacion_Actual', 'Estatus']], on="ID_Tarima", how="left")
            else:
                df_sub_det['Ubicacion_Actual'] = "Metales"
                df_sub_det['Estatus'] = "Disponible"

            # Rellenar nulos
            df_sub_det['Ubicacion_Actual'] = df_sub_det['Ubicacion_Actual'].fillna("Metales")
            df_sub_det['Estatus'] = df_sub_det['Estatus'].fillna("Disponible")
            df_sub_det['Cantidad'] = pd.to_numeric(df_sub_det['Cantidad'], errors='coerce').fillna(0).astype(int)

            # Cruzar con BD_Datos_Generales_Remision si está remesada
            rem_map = {}
            if not df_remisiones.empty:
                import ast
                for _, r_row in df_remisiones.iterrows():
                    fol = str(r_row.get('Folio_Remision', ''))
                    fec = str(r_row.get('Fecha_Hora_Salida', '')).split()[0]
                    rec = str(r_row.get('Nombre_Receptor', ''))
                    asoc = r_row.get('Tarimas_Asociadas', '')
                    if isinstance(asoc, str):
                        try: asoc = ast.literal_eval(asoc)
                        except Exception: asoc = [asoc]
                    if isinstance(asoc, list):
                        for t_id in asoc:
                            rem_map[str(t_id).strip()] = f"Remisión {fol} ({fec}) ➡️ {rec}"

            df_sub_det['Detalle_Remision'] = df_sub_det['ID_Tarima'].astype(str).str.strip().map(lambda x: rem_map.get(x, "N/A (En Almacén / Planta)"))

            # Métricas agregadas
            pzs_disponibles = df_sub_det[df_sub_det['Estatus'] == 'Disponible']['Cantidad'].sum()
            pzs_remesadas = df_sub_det[df_sub_det['Estatus'] == 'Remesada']['Cantidad'].sum()
            pzs_totales = df_sub_det['Cantidad'].sum()
            tarimas_totales = df_sub_det['ID_Tarima'].nunique()

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("📦 Piezas Disponibles (Planta)", f"{pzs_disponibles:,} PZS")
            m2.metric("🚚 Piezas Remesadas (Enviadas)", f"{pzs_remesadas:,} PZS")
            m3.metric("🏗️ Total Piezas Registradas", f"{pzs_totales:,} PZS")
            m4.metric("📊 Total Tarimas Físicas", f"{tarimas_totales} Tarimas")

            st.write("")
            st.markdown("#### 📍 Desglose de Tarimas y Ubicación Exacta")
            
            cols_mostrar = ['ID_Tarima', 'Ubicacion_Actual', 'Estatus', 'Cantidad', 'Proyecto', 'PO', 'Parcialidad', 'Descripcion', 'Detalle_Remision']
            df_tabla = df_sub_det[[c for c in cols_mostrar if c in df_sub_det.columns]].copy()
            df_tabla = df_tabla.rename(columns={
                'ID_Tarima': 'ID Tarima',
                'Ubicacion_Actual': 'Ubicación Actual',
                'Estatus': 'Estatus Tarima',
                'Cantidad': 'Cantidad (Pzs)',
                'Descripcion': 'Descripción Proyecto',
                'Detalle_Remision': 'Estatus de Salida / Remisión'
            })

            st.dataframe(df_tabla, use_container_width=True, hide_index=True)

            # Exportar consulta a Excel
            buf_xl = io.BytesIO()
            with pd.ExcelWriter(buf_xl, engine='openpyxl') as writer:
                df_tabla.to_excel(writer, index=False, sheet_name=f"Consulta_{sku_seleccionado}")
            buf_xl.seek(0)
            
            st.download_button(
                label=f"📥 Descargar Reporte de Consulta Excel ({sku_seleccionado}.xlsx)",
                data=buf_xl.getvalue(),
                file_name=f"Consulta_Inventario_{sku_seleccionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_consulta_sku"
            )
    else:
        st.info("No hay datos de detalles de tarimas cargados.")

elif proyecto_seleccionado:
    st.markdown(f"## 📌 Consulta por Proyecto: `{proyecto_seleccionado}`")
    if not df_detalle.empty:
        df_proj = df_detalle[df_detalle['Proyecto'].astype(str).str.strip() == proyecto_seleccionado].copy()
        if not df_proj.empty:
            if not df_tarimas.empty:
                df_proj = pd.merge(df_proj, df_tarimas[['ID_Tarima', 'Ubicacion_Actual', 'Estatus']], on="ID_Tarima", how="left")
            else:
                df_proj['Ubicacion_Actual'] = "Metales"
                df_proj['Estatus'] = "Disponible"

            df_proj['Cantidad'] = pd.to_numeric(df_proj['Cantidad'], errors='coerce').fillna(0).astype(int)

            tot_pzs = df_proj['Cantidad'].sum()
            tot_skus = df_proj['SKU'].nunique()
            tot_tar = df_proj['ID_Tarima'].nunique()

            m1, m2, m3 = st.columns(3)
            m1.metric("🏗️ Total Piezas del Proyecto", f"{tot_pzs:,} PZS")
            m2.metric("🏷️ SKUs Diferentes", f"{tot_skus} SKUs")
            m3.metric("📦 Total Tarimas", f"{tot_tar} Tarimas")

            st.write("")
            st.markdown("#### 📑 Partidas del Proyecto")
            st.dataframe(df_proj[['ID_Tarima', 'SKU', 'PO', 'Parcialidad', 'Cantidad', 'Ubicacion_Actual', 'Estatus']], use_container_width=True, hide_index=True)

elif po_seleccionada:
    st.markdown(f"## 📄 Consulta por Orden de Compra (PO): `{po_seleccionada}`")
    if not df_detalle.empty:
        df_po = df_detalle[df_detalle['PO'].astype(str).str.strip() == po_seleccionada].copy()
        if not df_po.empty:
            if not df_tarimas.empty:
                df_po = pd.merge(df_po, df_tarimas[['ID_Tarima', 'Ubicacion_Actual', 'Estatus']], on="ID_Tarima", how="left")
            else:
                df_po['Ubicacion_Actual'] = "Metales"
                df_po['Estatus'] = "Disponible"

            df_po['Cantidad'] = pd.to_numeric(df_po['Cantidad'], errors='coerce').fillna(0).astype(int)

            tot_pzs = df_po['Cantidad'].sum()
            tot_skus = df_po['SKU'].nunique()
            tot_tar = df_po['ID_Tarima'].nunique()

            m1, m2, m3 = st.columns(3)
            m1.metric("📄 Total Piezas de la PO", f"{tot_pzs:,} PZS")
            m2.metric("🏷️ SKUs Diferentes", f"{tot_skus} SKUs")
            m3.metric("📦 Total Tarimas", f"{tot_tar} Tarimas")

            st.write("")
            st.markdown("#### 📑 Partidas de la PO")
            st.dataframe(df_po[['ID_Tarima', 'SKU', 'Proyecto', 'Parcialidad', 'Cantidad', 'Ubicacion_Actual', 'Estatus']], use_container_width=True, hide_index=True)

else:
    st.info("💡 **Instrucciones:** Seleccione un **SKU**, **Proyecto** o **Orden de Compra (PO)** en la parte superior para consultar la foto del producto, disponibilidad y ubicación exacta.")

st.markdown("---")
st.markdown("<p style='text-align: center; color: #94A3B8; font-size: 12px;'>Industria SIGRAMA S.A. de C.V. — Portal Interno de Consulta Exclusiva de Inventario</p>", unsafe_allow_html=True)
