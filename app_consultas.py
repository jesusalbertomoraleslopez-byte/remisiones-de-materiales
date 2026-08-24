import streamlit as st
import os
import pandas as pd
import datetime
import io
import requests
import urllib.parse
from PIL import Image
import glob

# 1. CONFIGURACIÓN E INTERFAZ BASE RESPONSIVA (Portal de Consulta e Historial Interno)
st.set_page_config(
    page_title="Consulta de Inventario e Historial — SIGRAMA Metales",
    layout="wide",
    page_icon="🔍"
)

# REPOSITORIO DE DATOS
REPO_OWNER = "jesusalbertomoraleslopez-byte"
REPO_NAME = "remisiones-de-materiales"
BRANCH = "main"

# Helper de lectura segura de Secrets
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
@st.cache_data(ttl=120, max_entries=15)
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

# 3. HELPER ROBUSTO DE BÚSQUEDA Y DESCARGA DE IMÁGENES DE SKUS
@st.cache_data(ttl=300, max_entries=50)
def obtener_imagen_sku(sku):
    """
    Busca la imagen de un SKU de forma insensible a mayúsculas/minúsculas.
    1. Revisa la carpeta local `imagenes_articulos/`.
    2. Si no la encuentra, consulta la API de GitHub y descarga la imagen para visualizarla.
    """
    if not sku or str(sku).strip().upper() in ['', 'NAN', 'NONE', 'N/A']:
        return None

    sku_clean = str(sku).strip().upper()
    os.makedirs("imagenes_articulos", exist_ok=True)

    # 1. Escaneo local (insensible a mayúsculas/minúsculas)
    if os.path.exists("imagenes_articulos"):
        for f in os.listdir("imagenes_articulos"):
            f_upper = f.upper()
            if f_upper.startswith(f"{sku_clean}(") or f_upper.startswith(f"{sku_clean}."):
                local_path = os.path.join("imagenes_articulos", f)
                try:
                    return Image.open(local_path)
                except Exception:
                    pass

    # 2. Escaneo remoto en GitHub (si la app corre en Streamlit Cloud y no tiene el archivo en disco)
    token = obtener_secret("github_token")
    if token:
        try:
            url_list = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/imagenes_articulos?ref={BRANCH}"
            headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
            res = requests.get(url_list, headers=headers)
            if res.status_code == 200:
                for item in res.json():
                    name_upper = item["name"].upper()
                    if name_upper.startswith(f"{sku_clean}(") or name_upper.startswith(f"{sku_clean}."):
                        download_url = item.get("download_url")
                        if not download_url:
                            download_url = f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/{BRANCH}/imagenes_articulos/{urllib.parse.quote(item['name'])}"
                        
                        img_res = requests.get(download_url)
                        if img_res.status_code == 200:
                            local_save_path = os.path.join("imagenes_articulos", item["name"])
                            with open(local_save_path, "wb") as f_out:
                                f_out.write(img_res.content)
                            try:
                                return Image.open(local_save_path)
                            except Exception:
                                pass
        except Exception:
            pass

    return None

# Cargar Bases de Datos
df_articulos = cargar_excel_desde_github("BD_Articulos.xlsx")
df_skus_aut = cargar_excel_desde_github("BD_SKUs_Autorizados.xlsx")
df_tarimas = cargar_excel_desde_github("BD_Tarimas.xlsx")
df_detalle = cargar_excel_desde_github("BD_Detalle_Tarimas.xlsx")
df_remisiones = cargar_excel_desde_github("BD_Datos_Generales_Remision.xlsx")
df_actividad = cargar_excel_desde_github("BD_Actividad_Log.xlsx")

# --- ENCABEZADO PRINCIPAL ---
col_logo, col_title = st.columns([1, 4])
with col_logo:
    if os.path.exists("logo_sigrama.png"):
        st.image("logo_sigrama.png", use_container_width=True)
    else:
        st.title("🏭")
with col_title:
    st.markdown("""
    <div class="main-header">
        <h2 style="margin:0; color:#FFFFFF; font-size:24px;">🔍 PORTAL DE CONSULTA DE INVENTARIOS E HISTORIAL</h2>
        <p style="margin:4px 0 0 0; color:#94A3B8; font-size:14px;">
            Industria SIGRAMA S.A. de C.V. — Verificación en tiempo real de fotografías de SKUs, inventario actual y registro histórico de movimientos.
        </p>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# --- OBTENER LISTA COMPLETA DE SKUS ---
skus_set = set()
if not df_articulos.empty and 'SKU' in df_articulos.columns:
    skus_set.update(df_articulos['SKU'].dropna().astype(str).str.strip().str.upper())
if not df_skus_aut.empty and 'SKU' in df_skus_aut.columns:
    skus_set.update(df_skus_aut['SKU'].dropna().astype(str).str.strip().str.upper())
if not df_detalle.empty and 'SKU' in df_detalle.columns:
    skus_set.update(df_detalle['SKU'].dropna().astype(str).str.strip().str.upper())

lista_skus = sorted([s for s in skus_set if s not in ['', 'NAN', 'NONE', 'N/A']])

# --- PESTAÑAS PRINCIPALES ---
tab_consulta, tab_historial_global = st.tabs([
    "🔍 Consulta por SKU / Proyecto / PO",
    "📜 Historial Global de Movimientos del Sistema"
])

# =============================================================================
# PESTAÑA 1: CONSULTA POR SKU, PROYECTO O PO
# =============================================================================
with tab_consulta:
    st.markdown("### 🔎 Buscar Material")
    subtab_sku, subtab_proj, subtab_po = st.tabs(["🏷️ Búsqueda por SKU", "📌 Búsqueda por Proyecto", "📄 Búsqueda por PO"])

    sku_seleccionado = None
    proyecto_seleccionado = None
    po_seleccionada = None

    with subtab_sku:
        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            sku_input = st.selectbox(
                "Seleccione o escriba el Código de SKU:",
                options=["(Seleccione un SKU...)"] + lista_skus,
                key="lookup_sku_select_v2"
            )
            if sku_input != "(Seleccione un SKU...)":
                sku_seleccionado = sku_input

        with col_s2:
            sku_manual = st.text_input("O escriba código manualmente:", placeholder="Ej: 11-A-6815-10", key="lookup_sku_text_v2")
            if sku_manual.strip():
                sku_seleccionado = sku_manual.strip().upper()

    with subtab_proj:
        if not df_detalle.empty and 'Proyecto' in df_detalle.columns:
            proyectos_disponibles = sorted([str(p).strip() for p in df_detalle['Proyecto'].dropna().unique() if str(p).strip() not in ['', 'nan', 'None']])
            proyecto_seleccionado = st.selectbox("Seleccione el Proyecto:", options=["(Seleccione un Proyecto...)"] + proyectos_disponibles, key="lookup_proj_select_v2")
            if proyecto_seleccionado == "(Seleccione un Proyecto...)":
                proyecto_seleccionado = None

    with subtab_po:
        if not df_detalle.empty and 'PO' in df_detalle.columns:
            pos_disponibles = sorted([str(p).strip() for p in df_detalle['PO'].dropna().unique() if str(p).strip() not in ['', 'nan', 'None']])
            po_seleccionada = st.selectbox("Seleccione la Orden de Compra (PO):", options=["(Seleccione una PO...)"] + pos_disponibles, key="lookup_po_select_v2")
            if po_seleccionada == "(Seleccione una PO...)":
                po_seleccionada = None

    st.write("---")

    # --- DESPLIEGUE PARA SKU SELECCIONADO ---
    if sku_seleccionado:
        st.markdown(f"## 📦 Resultado para el SKU: `{sku_seleccionado}`")
        
        col_info, col_img = st.columns([2, 1])
        
        # 1. FICHA TÉCNICA DEL ARTÍCULO
        with col_info:
            st.markdown("#### 📋 Ficha Técnica del Artículo")
            
            nombre_art = "No especificado en catálogo"
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

        # 2. FOTOGRAFÍA DEL ARTÍCULO (DISCO LOCAL + GITHUB API FALLBACK)
        with col_img:
            st.markdown("#### 🖼️ Fotografía del Producto")
            imagen_sku = obtener_imagen_sku(sku_seleccionado)
            if imagen_sku:
                st.image(imagen_sku, caption=f"Fotografía oficial de {sku_seleccionado}", use_container_width=True)
            else:
                st.info("📷 *Sin fotografía asignada en catálogo.*")

        st.write("---")

        # 3. DETALLE DE INVENTARIO Y UBICACIÓN FÍSICA
        if not df_detalle.empty:
            df_sub_det = df_detalle[df_detalle['SKU'].astype(str).str.strip().str.upper() == sku_seleccionado].copy()
            
            if df_sub_det.empty:
                st.info(f"ℹ️ **Estado de Registro:** El SKU **{sku_seleccionado}** está autorizado en catálogo master, pero **aún no cuenta con paquetes o tarimas creadas en planta**.")
            else:
                # Cruzar con BD_Tarimas para obtener Fecha_Creacion, Creado_Por, Ubicacion_Actual y Estatus
                if not df_tarimas.empty:
                    df_sub_det = pd.merge(
                        df_sub_det, 
                        df_tarimas[['ID_Tarima', 'Fecha_Creacion', 'Creado_Por', 'Ubicacion_Actual', 'Estatus']], 
                        on="ID_Tarima", 
                        how="left"
                    )
                else:
                    df_sub_det['Fecha_Creacion'] = "N/A"
                    df_sub_det['Creado_Por'] = "N/A"
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
                        dir_rec = str(r_row.get('Direccion_Receptor', ''))
                        asoc = r_row.get('Tarimas_Asociadas', '')
                        if isinstance(asoc, str):
                            try: asoc = ast.literal_eval(asoc)
                            except Exception: asoc = [asoc]
                        if isinstance(asoc, list):
                            for t_id in asoc:
                                rem_map[str(t_id).strip()] = f"Remisión {fol} ({fec}) ➡️ {rec} [{dir_rec}]"

                df_sub_det['Detalle_Remision'] = df_sub_det['ID_Tarima'].astype(str).str.strip().map(lambda x: rem_map.get(x, "En Planta / Almacén"))

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
                st.markdown("#### 📍 Desglose Histórico de Tarimas y Ubicación Exacta")
                
                cols_mostrar = ['ID_Tarima', 'Fecha_Creacion', 'Creado_Por', 'Ubicacion_Actual', 'Estatus', 'Cantidad', 'Proyecto', 'PO', 'Parcialidad', 'Descripcion', 'Detalle_Remision']
                df_tabla = df_sub_det[[c for c in cols_mostrar if c in df_sub_det.columns]].copy()
                df_tabla = df_tabla.rename(columns={
                    'ID_Tarima': 'ID Tarima (TPM)',
                    'Fecha_Creacion': 'Fecha Empaque',
                    'Creado_Por': 'Líder Empaque',
                    'Ubicacion_Actual': 'Ubicación Actual',
                    'Estatus': 'Estatus Tarima',
                    'Cantidad': 'Piezas',
                    'Descripcion': 'Descripción Proyecto',
                    'Detalle_Remision': 'Estatus de Remisión / Destino'
                })

                st.dataframe(df_tabla, use_container_width=True, hide_index=True)

                # Exportar consulta a Excel
                buf_xl = io.BytesIO()
                with pd.ExcelWriter(buf_xl, engine='openpyxl') as writer:
                    df_tabla.to_excel(writer, index=False, sheet_name=f"Consulta_{sku_seleccionado}")
                buf_xl.seek(0)
                
                st.download_button(
                    label=f"📥 Descargar Registro Histórico en Excel ({sku_seleccionado}.xlsx)",
                    data=buf_xl.getvalue(),
                    file_name=f"Consulta_Inventario_{sku_seleccionado}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_consulta_sku_v2"
                )

        # 4. BUSCAR EN LOG DE ACTIVIDAD DEL SISTEMA
        if not df_actividad.empty:
            st.write("")
            st.markdown(f"#### 📜 Eventos e Historial de Bitácora Registrados para `{sku_seleccionado}`")
            
            mask_act = df_actividad['Descripcion'].astype(str).str.contains(sku_seleccionado, case=False, na=False) | \
                       df_actividad['ID_Referencia'].astype(str).str.contains(sku_seleccionado, case=False, na=False)
            df_act_sku = df_actividad[mask_act].copy()
            
            if not df_act_sku.empty:
                st.dataframe(df_act_sku, use_container_width=True, hide_index=True)
            else:
                st.caption("No se encontraron registros adicionales en la bitácora log para este SKU.")

    elif proyecto_seleccionado:
        st.markdown(f"## 📌 Consulta por Proyecto: `{proyecto_seleccionado}`")
        if not df_detalle.empty:
            df_proj = df_detalle[df_detalle['Proyecto'].astype(str).str.strip() == proyecto_seleccionado].copy()
            if not df_proj.empty:
                if not df_tarimas.empty:
                    df_proj = pd.merge(df_proj, df_tarimas[['ID_Tarima', 'Fecha_Creacion', 'Creado_Por', 'Ubicacion_Actual', 'Estatus']], on="ID_Tarima", how="left")
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
                st.dataframe(df_proj[['ID_Tarima', 'Fecha_Creacion', 'SKU', 'PO', 'Parcialidad', 'Cantidad', 'Ubicacion_Actual', 'Estatus']], use_container_width=True, hide_index=True)

    elif po_seleccionada:
        st.markdown(f"## 📄 Consulta por Orden de Compra (PO): `{po_seleccionada}`")
        if not df_detalle.empty:
            df_po = df_detalle[df_detalle['PO'].astype(str).str.strip() == po_seleccionada].copy()
            if not df_po.empty:
                if not df_tarimas.empty:
                    df_po = pd.merge(df_po, df_tarimas[['ID_Tarima', 'Fecha_Creacion', 'Creado_Por', 'Ubicacion_Actual', 'Estatus']], on="ID_Tarima", how="left")
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
                st.dataframe(df_po[['ID_Tarima', 'Fecha_Creacion', 'SKU', 'Proyecto', 'Parcialidad', 'Cantidad', 'Ubicacion_Actual', 'Estatus']], use_container_width=True, hide_index=True)

    else:
        st.info("💡 **Instrucciones:** Seleccione un **SKU**, **Proyecto** o **Orden de Compra (PO)** para consultar la foto del producto, disponibilidad, ubicación física e historial de movimientos.")

# =============================================================================
# PESTAÑA 2: HISTORIAL GLOBAL DE MOVIMIENTOS Y EVENTOS
# =============================================================================
with tab_historial_global:
    st.markdown("### 📜 Bitácora e Historial Global de Movimientos")
    st.markdown("Consulte el registro de todas las transacciones realizadas en la plataforma (Empaques, Remisiones, Modificaciones y Entregas):")
    
    if not df_actividad.empty:
        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            query_search = st.text_input("🔍 Filtrar por palabra clave (SKU, TPM-XXXX, Folio Remisión, Usuario, PO, Cliente):", key="search_log_global")
        with col_f2:
            tipos_evento = ["Todos"] + sorted(df_actividad['Tipo_Evento'].dropna().astype(str).unique().tolist())
            filtro_tipo = st.selectbox("Filtrar por Tipo de Evento:", options=tipos_evento, key="filter_event_type_global")

        df_log_filt = df_actividad.copy()

        if filtro_tipo != "Todos":
            df_log_filt = df_log_filt[df_log_filt['Tipo_Evento'] == filtro_tipo]

        if query_search.strip():
            q = query_search.strip().lower()
            mask = df_log_filt.apply(lambda row: row.astype(str).str.lower().str.contains(q).any(), axis=1)
            df_log_filt = df_log_filt[mask]

        st.write("")
        st.write(f"📊 **Total de registros encontrados:** {len(df_log_filt):,}")
        st.dataframe(df_log_filt, use_container_width=True, hide_index=True)

        buf_log_xl = io.BytesIO()
        with pd.ExcelWriter(buf_log_xl, engine='openpyxl') as writer:
            df_log_filt.to_excel(writer, index=False, sheet_name="Historial_Movimientos")
        buf_log_xl.seek(0)
        
        st.download_button(
            label="📥 Descargar Historial Global en Excel (.xlsx)",
            data=buf_log_xl.getvalue(),
            file_name="Historial_Movimientos_Sigrama.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_dl_log_global"
        )
    else:
        st.warning("No se encontraron registros de historial en la base de datos `BD_Actividad_Log.xlsx`.")

st.markdown("---")
st.markdown("<p style='text-align: center; color: #94A3B8; font-size: 12px;'>Industria SIGRAMA S.A. de C.V. — Portal Interno de Consulta Exclusiva e Historial</p>", unsafe_allow_html=True)
