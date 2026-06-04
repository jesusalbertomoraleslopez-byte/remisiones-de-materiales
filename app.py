import streamlit as st
import pandas as pd
import datetime
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# 1. CONFIGURACIÓN DE LA PÁGINA WEB
st.set_page_config(page_title="Remisiones de Materiales", layout="wide", page_icon="📦")

# 2. INICIALIZACIÓN DE MODELOS RELACIONALES EN MEMORIA (PANDAS)
if "BD_Articulos" not in st.session_state:
    st.session_state.BD_Articulos = pd.DataFrame([
        {"SKU": "SKU-001", "Nombre": "Lámina de Acero Galvanizado", "Calibre_Espesor": "Calibre 22", "Dimensiones_Pieza": "3x10 ft", "Acabado_Superficial": "Zintro"},
        {"SKU": "SKU-002", "Nombre": "Placa de Acero Comercial", "Calibre_Espesor": "1/4 pulgada", "Dimensiones_Pieza": "4x8 ft", "Acabado_Superficial": "Negro"},
        {"SKU": "SKU-003", "Nombre": "Perfil Tubular Cuadrado", "Calibre_Espesor": "Calibre 14", "Dimensiones_Pieza": "2x2 pulg x 6m", "Acabado_Superficial": "Pintado"}
    ])

if "BD_Ordenes_Compra" not in st.session_state:
    st.session_state.BD_Ordenes_Compra = pd.DataFrame([
        {"PO": "PO-10001", "SKU": "SKU-001", "Cantidad_Solicitada": 500},
        {"PO": "PO-10001", "SKU": "SKU-002", "Cantidad_Solicitada": 200},
        {"PO": "PO-10002", "SKU": "SKU-003", "Cantidad_Solicitada": 1000}
    ])

if "BD_Tarimas" not in st.session_state:
    st.session_state.BD_Tarimas = pd.DataFrame(columns=[
        "ID_Tarima", "Tarima_Origen_Excel", "Fecha_Creacion", "Ubicacion_Actual", "Creado_Por", "Tipo_Tarima", "Estatus"
    ])

if "BD_Detalle_Tarimas" not in st.session_state:
    st.session_state.BD_Detalle_Tarimas = pd.DataFrame(columns=["ID_Detalle", "ID_Tarima", "SKU", "PO", "Cantidad"])

if "BD_Datos_Generales_Remision" not in st.session_state:
    st.session_state.BD_Datos_Generales_Remision = pd.DataFrame(columns=[
        "ID_Remision", "Folio_Remision", "Fecha_Hora_Salida", "Nombre_Emisor", "Direccion_Emisor", "Nombre_Receptor", "Direccion_Receptor", "Tarimas_Asociadas"
    ])

# 3. CONTROL DE SEGURIDAD MEDIANTE ST.SECRETS
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

# MENÚ LATERAL DE SECCIONES
st.sidebar.title("🧭 Navegación")
opcion_menu = st.sidebar.radio(
    "Seleccione un Módulo:",
    ["📊 Dashboard e Históricos", "🔍 Centro de Consultas", "📦 Módulo Tarimas", "🚚 Módulo Remisiones"]
)
# 4. MOTOR DE RENDERIZADO PARA REPORTLAB
def generar_pdf_tarima(id_tarima):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story, styles = [], getSampleStyleSheet()
    
    tarima_info = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['ID_Tarima'] == id_tarima].iloc[0]
    detalles = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'] == id_tarima]
    
    # Hoja 1: Número de tarima gigante y centrado
    style_g = ParagraphStyle('G', parent=styles['Heading1'], fontSize=54, leading=60, alignment=1)
    story.append(Spacer(1, 2 * inch))
    story.append(Paragraph(f"TARIMA<br/><br/><b>#{id_tarima}</b>", style_g))
    story.append(PageBreak())
    
    # Hoja 2: Datos de creación y Cantidad Gigante Negrita
    style_n = styles['Normal']
    style_ng = ParagraphStyle('NG', parent=styles['Heading2'], fontSize=28, leading=34, alignment=1)
    story.append(Paragraph(f"<b>Detalle de Control Interno - Tarima #{id_tarima}</b>", styles['Heading2']))
    story.append(Paragraph(f"<b>Creado por:</b> {tarima_info['Creado_Por']} | <b>Fecha:</b> {tarima_info['Fecha_Creacion']}", style_n))
    story.append(Spacer(1, 0.3 * inch))
    
    for _, item in detalles.iterrows():
        art = st.session_state.BD_Articulos[st.session_state.BD_Articulos['SKU'] == item['SKU']]
        nom_art = art.iloc[0]['Nombre'] if not art.empty else "Desconocido"
        story.append(Paragraph(f"<b>PO Asociada:</b> {item['PO']} | <b>SKU:</b> {item['SKU']}", style_n))
        story.append(Paragraph(f"<b>Producto:</b> {nom_art}", style_n))
        story.append(Spacer(1, 0.4 * inch))
        story.append(Paragraph(f"CANTIDAD TOTAL:", style_n))
        story.append(Paragraph(f"<b>{int(item['Cantidad'])} PZS</b>", style_ng))
        
    doc.build(story)
    buffer.seek(0)
    return buffer

def generar_pdf_remision_general(datos_remision, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story, styles = [], getSampleStyleSheet()
    
    story.append(Paragraph(f"<b>REMISIÓN DE MATERIALES: {datos_remision['Folio_Remision']}</b>", styles['Title']))
    story.append(Spacer(1, 0.2 * inch))
    
    data_e_r = [
        [Paragraph(f"<b>EMISOR:</b> {datos_remision['Nombre_Emisor']}", styles['Normal']), 
         Paragraph(f"<b>RECEPTOR:</b> {datos_remision['Nombre_Receptor']}", styles['Normal'])],
        [Paragraph(f"<b>Dirección:</b> {datos_remision['Direccion_Emisor']}", styles['Normal']), 
         Paragraph(f"<b>Dirección:</b> {datos_remision['Direccion_Receptor']}", styles['Normal'])]
    ]
    t_info = Table(data_e_r, colWidths=[3.25 * inch, 3.25 * inch])
    story.append(t_info)
    story.append(Spacer(1, 0.4 * inch))
    
    tabla_data = [["Tarima", "PO", "SKU", "Cantidad"]]
    for _, row in df_detalles_remision.iterrows():
        tabla_data.append([str(row['ID_Tarima']), str(row['PO']), str(row['SKU']), str(int(row['Cantidad']))])
        
    t_prod = Table(tabla_data, colWidths=[1.5 * inch, 1.5 * inch, 2.0 * inch, 1.5 * inch])
    t_prod.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    story.append(t_prod)
    doc.build(story)
    buffer.seek(0)
    return buffer

def generar_pdf_anexo_tarimas(lista_tarimas_id, df_detalles_remision):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story, styles = [], getSampleStyleSheet()
    
    for t_id in lista_tarimas_id:
        # Hoja 1: Carátula Anexo
        style_c = ParagraphStyle('C', parent=styles['Heading1'], fontSize=42, alignment=1)
        story.append(Spacer(1, 2 * inch))
        story.append(Paragraph(f"ANEXO: TARIMA #{t_id}", style_c))
        story.append(Spacer(1, 0.5 * inch))
        t_bar = Table([["|||||||||||||||||||||||||||||||"], [f"*TARIMA-{t_id}*"]], colWidths=[6.5 * inch])
        t_bar.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('TEXTCOLOR', (0,0), (-1,1), colors.darkgray)]))
        story.append(t_bar)
        story.append(PageBreak())
        
        # Hoja 2: Detalle con Sombreado Gris Tenue
        story.append(Paragraph(f"<b>DETALLE ESPECÍFICO - TARIMA {t_id}</b>", styles['Heading2']))
        sub_det = df_detalles_remision[df_detalles_remision['ID_Tarima'] == t_id]
        tabla_anexo = [["PO", "SKU", "Cantidad"]]
        for _, s_row in sub_det.iterrows():
            tabla_anexo.append([str(s_row['PO']), str(s_row['SKU']), str(int(s_row['Cantidad']))])
            
        t_det = Table(tabla_anexo, colWidths=[2.1 * inch, 2.2 * inch, 2.2 * inch])
        t_det.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkgray),
            ('BACKGROUND', (0,1), (-1,-1), colors.lightgrey), # Fondo contrastante
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 1, colors.white)
        ]))
        story.append(t_det)
        story.append(PageBreak())
        
    if story: story.pop()
    doc.build(story)
    buffer.seek(0)
    return buffer

# ==========================================
# MÓDULOS DE VISUALIZACIÓN ACCESIBLES
# ==========================================
if opcion_menu == "📊 Dashboard e Históricos":
    st.title("📊 Dashboard General de Operaciones")
    col_f1, col_f2 = st.columns(2)
    with col_f1: f_inicio = st.date_input("Fecha Inicial", datetime.date.today() - datetime.timedelta(days=7))
    with col_f2: f_fin = st.date_input("Fecha Final", datetime.date.today())
        
    total_tarimas = len(st.session_state.BD_Tarimas)
    disponibles = len(st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible'])
    remesadas = len(st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Remesada'])
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Tarimas Registradas", total_tarimas)
    m2.metric("🟢 Tarimas Disponibles", disponibles)
    m3.metric("🚚 Tarimas Remesadas", remesadas)
    
    st.dataframe(st.session_state.BD_Tarimas, use_container_width=True)
    if st.button("🖨️ Generar PDF del Dashboard del Día"):
        st.info("Reporte preliminar del día estructurado con éxito.")

elif opcion_menu == "🔍 Centro de Consultas":
    st.title("🔍 Centro de Consultas Avanzado")
    col_sel, col_val = st.columns(2)
    with col_sel: tipo_filtro = st.selectbox("Filtrar por:", ["Ninguno", "SKU", "PO", "Folio_Remision"])
    with col_val: valor_filtro = st.text_input("Ingrese término de búsqueda:")
        
    resultado_vista = st.session_state.BD_Detalle_Tarimas.copy()
    if tipo_filtro == "SKU" and valor_filtro:
        resultado_vista = resultado_vista[resultado_vista['SKU'].str.contains(valor_filtro, case=False)]
    elif tipo_filtro == "PO" and valor_filtro:
        resultado_vista = resultado_vista[resultado_vista['PO'].str.contains(valor_filtro, case=False)]
    elif tipo_filtro == "Folio_Remision" and valor_filtro:
        remisiones_match = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].str.contains(valor_filtro, case=False)]
        tarimas_validas = [t for r in remisiones_match['Tarimas_Asociadas'] for t in r]
        resultado_vista = resultado_vista[resultado_vista['ID_Tarima'].isin(tarimas_validas)]

    st.dataframe(resultado_vista, use_container_width=True)
    
    if not resultado_vista.empty:
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
            resultado_vista.to_excel(writer, index=False, sheet_name='Reporte')
        st.download_button(
            label="📥 Descargar Reporte en Excel (.xlsx)",
            data=towrite.getvalue(),
            file_name=f"Reporte_Consolidado_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
# ==========================================
# MÓDULOS DE OPERACIÓN PROTEGIDOS
# ==========================================
elif opcion_menu == "📦 Módulo Tarimas":
    st.title("📦 Carga y Administración de Tarimas")
    
    # BOTÓN DE DESCARGA DE PLANTILLA MUESTRA DE EJEMPLO
    st.subheader("📋 Formato Requerido")
    st.write("Antes de subir tus datos, descarga y utiliza esta plantilla de ejemplo:")
    
    df_plantilla_ejemplo = pd.DataFrame([
        {"Tarima": "Bulto_Ejemplo_1", "Producto/SKU": "SKU-001", "PO": "PO-10001", "Cantidad": 150},
        {"Tarima": "Bulto_Ejemplo_1", "Producto/SKU": "SKU-002", "PO": "PO-10001", "Cantidad": 80},
        {"Tarima": "Bulto_Ejemplo_2", "Producto/SKU": "SKU-001", "PO": "PO-10001", "Cantidad": 200},
        {"Tarima": "Bulto_Ejemplo_3", "Producto/SKU": "SKU-003", "PO": "PO-10002", "Cantidad": 500}
    ])
    
    buffer_plantilla = io.BytesIO()
    with pd.ExcelWriter(buffer_plantilla, engine='openpyxl') as writer_p:
        df_plantilla_ejemplo.to_excel(writer_p, index=False, sheet_name='Plantilla_Tarimas')
    buffer_plantilla.seek(0)
    
    st.download_button(
        label="📥 Descargar Plantilla de Ejemplo (.xlsx)",
        data=buffer_plantilla.getvalue(),
        file_name="plantilla_remisiones_tarimas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.write("---")
    
    if not is_admin:
        st.error("🔒 Área Bloqueada: Se requiere contraseña de Administrador para subir datos.")
    else:
        st.success("🔓 Acceso Autorizado para Operaciones.")
        archivo_cargado = st.file_uploader("Suba la Plantilla Excel de Tarimas rellenada", type=["xlsx"])
        col_t1, col_t2 = st.columns(2)
        with col_t1: tipo_t = st.selectbox("Tipo de Tarima:", ["Cuadrada", "Rectangular"])
        with col_t2: operador = st.text_input("Operador a Cargo:", "Supervisor General")
            
        if archivo_cargado and st.button("Procesar e Integrar Plantilla"):
            try:
                df_excel = pd.read_excel(archivo_cargado)
                columnas_requeridas = ["Tarima", "Producto/SKU", "PO", "Cantidad"]
                if not all(col in df_excel.columns for col in columnas_requeridas):
                    st.error("❌ Error: Las columnas del archivo no coinciden con la plantilla permitida.")
                else:
                    for t_orig in df_excel['Tarima'].unique():
                        nuevo_id = len(st.session_state.BD_Tarimas) + 1
                        nueva_t = {"ID_Tarima": nuevo_id, "Tarima_Origen_Excel": t_orig, "Fecha_Creacion": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "Ubicacion_Actual": "Almacén Central", "Creado_Por": operador, "Tipo_Tarima": tipo_t, "Estatus": "Disponible"}
                        st.session_state.BD_Tarimas = pd.concat([st.session_state.BD_Tarimas, pd.DataFrame([nueva_t])], ignore_index=True)
                        
                        items = df_excel[df_excel['Tarima'] == t_orig]
                        for _, item in items.iterrows():
                            nuevo_det = {"ID_Detalle": len(st.session_state.BD_Detalle_Tarimas)+1, "ID_Tarima": nuevo_id, "SKU": item['Producto/SKU'], "PO": item['PO'], "Cantidad": item['Cantidad']}
                            st.session_state.BD_Detalle_Tarimas = pd.concat([st.session_state.BD_Detalle_Tarimas, pd.DataFrame([nuevo_det])], ignore_index=True)
                    st.success("¡Plantilla integrada correctamente!")
            except Exception as e: st.error(f"Error en procesamiento: {e}")
                
    st.write("---")
    st.subheader("🖨️ Impresión Individual de Tarimas")
    if not st.session_state.BD_Tarimas.empty:
        tarima_imprimir = st.selectbox("Seleccione Tarima a Imprimir:", st.session_state.BD_Tarimas['ID_Tarima'].unique())
        if st.checkbox("Preparar Archivo de Impresión"):
            st.download_button(label=f"📥 Descargar PDF Tarima #{tarima_imprimir}", data=generar_pdf_tarima(tarima_imprimir), file_name=f"Tarima_{tarima_imprimir}.pdf", mime="application/pdf")

elif opcion_menu == "🚚 Módulo Remisiones":
    st.title("🚚 Generación de Remisiones de Salida")
    tarimas_disponibles = st.session_state.BD_Tarimas[st.session_state.BD_Tarimas['Estatus'] == 'Disponible']['ID_Tarima'].tolist()
    
    if not tarimas_disponibles:
        st.warning("⚠️ No existen tarimas bajo estatus 'Disponible' para asignación.")
    else:
        tarimas_seleccionadas = st.multiselect("Seleccione Tarimas Disponibles:", options=tarimas_disponibles)
        col_e, col_r = st.columns(2)
        with col_e:
            nom_emisor = st.text_input("Emisor:", "Planta Industrial Centro")
            dir_emisor = st.text_input("Dirección Emisor:", "Av. de la Industria #450")
        with col_r:
            nom_receptor = st.text_input("Receptor / Cliente:")
            dir_receptor = st.text_input("Dirección Destino:")
            
        if not is_admin:
            st.error("🔒 Operación Bloqueada: Se requiere contraseña de Administrador para guardar cambios.")
        else:
            if st.button("🚀 Confirmar Salida y Generar Nueva Remisión"):
                if not tarimas_seleccionadas or not nom_receptor: st.error("Por favor complete los campos obligatorios.")
                else:
                    n_id = len(st.session_state.BD_Datos_Generales_Remision) + 1
                    folio = f"REM-{datetime.date.today().year}-{1000 + n_id}"
                    reg = {"ID_Remision": n_id, "Folio_Remision": folio, "Fecha_Hora_Salida": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "Nombre_Emisor": nom_emisor, "Direccion_Emisor": dir_emisor, "Nombre_Receptor": nom_receptor, "Direccion_Receptor": dir_receptor, "Tarimas_Asociadas": tarimas_seleccionadas}
                    
                    st.session_state.BD_Datos_Generales_Remision = pd.concat([st.session_state.BD_Datos_Generales_Remision, pd.DataFrame([reg])], ignore_index=True)
                    st.session_state.BD_Tarimas.loc[st.session_state.BD_Tarimas['ID_Tarima'].isin(tarimas_seleccionadas), 'Estatus'] = 'Remesada'
                    st.success(f"✅ ¡Remisión {folio} Generada Exitosamente!")
                    
    st.write("---")
    st.subheader("🖨️ Descarga Documental de Remisiones")
    if not st.session_state.BD_Datos_Generales_Remision.empty:
        rem_sel = st.selectbox("Seleccione Folio para Descarga:", st.session_state.BD_Datos_Generales_Remision['Folio_Remision'].unique())
        row = st.session_state.BD_Datos_Generales_Remision[st.session_state.BD_Datos_Generales_Remision['Folio_Remision'] == rem_sel].iloc[0]
        df_det = st.session_state.BD_Detalle_Tarimas[st.session_state.BD_Detalle_Tarimas['ID_Tarima'].isin(row['Tarimas_Asociadas'])]
        
        c1, c2 = st.columns(2)
        with c1: st.download_button("📥 Descargar Remisión (PDF)", data=generar_pdf_remision_general(row, df_det), file_name=f"Remision_{rem_sel}.pdf")
        with c2: st.download_button("📥 Descargar Anexo Tarimas (PDF)", data=generar_pdf_anexo_tarimas(row['Tarimas_Asociadas'], df_det), file_name=f"Anexo_{rem_sel}.pdf")
