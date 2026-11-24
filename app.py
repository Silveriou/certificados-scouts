import os
import json
import base64
import importlib
import streamlit as st
from datetime import datetime, date
import generator
importlib.reload(generator)
from generator import generar_certificado_pdf, pdf_to_image_bytes, DEFAULT_CONFIG

# Rutas base
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
PLANTILLAS_DIR = os.path.join(BASE_DIR, "plantillas")
FIRMAS_DIR = os.path.join(BASE_DIR, "firmas")
DATOS_DIR = os.path.join(BASE_DIR, "datos")
GENERADOS_DIR = os.path.join(BASE_DIR, "generados")

for d in [ASSETS_DIR, PLANTILLAS_DIR, FIRMAS_DIR, DATOS_DIR, GENERADOS_DIR]:
    os.makedirs(d, exist_ok=True)

LOGO_SCOUT_PATH = os.path.join(ASSETS_DIR, "logo_scout.png")

# Configuración de página con el logo scout oficial
st.set_page_config(
    page_title="ASV | Condecoraciones Scouts",
    page_icon=LOGO_SCOUT_PATH if os.path.exists(LOGO_SCOUT_PATH) else "⚜️",
    layout="wide"
)

# Nombres de meses en español
MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

FIRMANTES_FILE = os.path.join(DATOS_DIR, "firmantes.json")
CONDECORACIONES_FILE = os.path.join(DATOS_DIR, "condecoraciones.json")
CODIGOS_FILE = os.path.join(DATOS_DIR, "codigos.json")

def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

def cargar_firmantes():
    if os.path.exists(FIRMANTES_FILE):
        with open(FIRMANTES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def guardar_firmantes(firmantes):
    with open(FIRMANTES_FILE, "w", encoding="utf-8") as f:
        json.dump(firmantes, f, indent=2, ensure_ascii=False)

def cargar_condecoraciones():
    if os.path.exists(CONDECORACIONES_FILE):
        with open(CONDECORACIONES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            for item in data:
                if "config" not in item:
                    item["config"] = DEFAULT_CONFIG.copy()
            return data
    return []

def guardar_condecoraciones(condec):
    with open(CONDECORACIONES_FILE, "w", encoding="utf-8") as f:
        json.dump(condec, f, indent=2, ensure_ascii=False)

def cargar_codigos():
    if os.path.exists(CODIGOS_FILE):
        with open(CODIGOS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def guardar_codigos(codigos):
    with open(CODIGOS_FILE, "w", encoding="utf-8") as f:
        json.dump(codigos, f, indent=2, ensure_ascii=False)

# Estilos personalizados Scouts
logo_b64 = get_base64_image(LOGO_SCOUT_PATH)
st.markdown(f"""
<style>
/* Botones primarios en Morado Scout Oficial (#622599) */
.stButton>button {{
    background-color: #622599 !important;
    color: #FFFFFF !important;
    border-radius: 8px !important;
    border: none !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.2rem !important;
    transition: all 0.25s ease-in-out !important;
}}
.stButton>button:hover {{
    background-color: #4B1779 !important;
    color: #FFFFFF !important;
    box-shadow: 0 4px 14px rgba(98, 37, 153, 0.4) !important;
    transform: translateY(-1px);
}}
.stDownloadButton>button {{
    background: linear-gradient(135deg, #622599 0%, #7E36C2 100%) !important;
    color: #FFFFFF !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 0.65rem 1.2rem !important;
    box-shadow: 0 4px 14px rgba(98, 37, 153, 0.35) !important;
    border: none !important;
    transition: all 0.25s ease-in-out !important;
}}
.stDownloadButton>button:hover {{
    background: linear-gradient(135deg, #4B1779 0%, #622599 100%) !important;
    box-shadow: 0 6px 18px rgba(98, 37, 153, 0.5) !important;
    color: #FFFFFF !important;
    transform: translateY(-1px);
}}
/* Encabezado Scout Oficial */
.scout-header {{
    display: flex;
    align-items: center;
    gap: 22px;
    padding: 20px 26px;
    background: linear-gradient(135deg, #622599 0%, #401066 100%);
    border-radius: 14px;
    margin-bottom: 25px;
    box-shadow: 0 8px 24px rgba(98, 37, 153, 0.28);
}}
.scout-header-img {{
    width: 80px;
    height: 80px;
    border-radius: 50%;
    background: #FFFFFF;
    padding: 4px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    flex-shrink: 0;
}}
.scout-header-text h1 {{
    color: #FFFFFF !important;
    margin: 0 !important;
    font-size: 1.85rem !important;
    font-weight: 800 !important;
    letter-spacing: -0.5px;
}}
.scout-header-text p {{
    color: #EBDCF9 !important;
    margin: 4px 0 0 0 !important;
    font-size: 1rem !important;
}}
.scout-badge {{
    background: #EFE7F8;
    color: #4B1779;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 700;
    display: inline-block;
    border: 1px solid #D5BEF0;
    margin-top: 6px;
}}
/* Pie de página */
.scout-footer {{
    text-align: center;
    color: #6C5B80;
    font-size: 0.88rem;
    padding: 30px 0 10px 0;
    border-top: 1px solid #EFE7F8;
    margin-top: 45px;
}}
</style>

<div class="scout-header">
    <img src="data:image/png;base64,{logo_b64}" class="scout-header-img" alt="Flor de Lis Scout">
    <div class="scout-header-text">
        <h1>Asociación de Scouts de Venezuela</h1>
        <p>Sistema Oficial de Emisión de Condecoraciones y Certificados</p>
        <span class="scout-badge">⚜️ Siempre Listos para Servir • Región Zulia</span>
    </div>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "📜 Emitir Certificado",
    "🏷️ Códigos de Control",
    "✍️ Gestión de Firmantes",
    "📁 Formatos y Plantillas"
])

# -------------------------------------------------------------------------
# TAB 1: EMITIR CERTIFICADO
# -------------------------------------------------------------------------
with tab1:
    condecoraciones = cargar_condecoraciones()
    firmantes = cargar_firmantes()
    codigos_guardados = cargar_codigos()
    
    if not condecoraciones:
        st.warning("No hay formatos de condecoración cargados. Por favor ve a la pestaña 'Formatos y Plantillas' para agregar uno.")
    elif not firmantes:
        st.warning("No hay firmantes registrados. Por favor ve a la pestaña 'Gestión de Firmantes' para agregar al menos uno.")
    else:
        col_form, col_prev = st.columns([1, 1], gap="large")
        
        with col_form:
            st.subheader("1. Condecoración y Galardonado")
            
            nombres_condec = [c["nombre"] for c in condecoraciones]
            condec_seleccionada = st.selectbox("Seleccione la Condecoración / Formato:", nombres_condec)
            condec_obj = next((c for c in condecoraciones if c["nombre"] == condec_seleccionada), None)
            
            nombre_galardonado = st.text_input(
                "Nombre y Apellido:",
                value="SILVERIO A. URDANETA S.",
                help="Se imprimirá en mayúsculas ocupando aproximadamente el 90% de la línea oficial"
            )
            
            cedula = st.text_input(
                "Cédula de Identidad:",
                value="11.289.602",
                help="Coloque el número (ej: 11.289.602). El prefijo C.I. V- ya está impreso en el formato."
            )
            
            st.markdown("---")
            st.subheader("2. Fecha y Código de Registro")
            
            col_fec_pick, col_fec_txt = st.columns(2)
            with col_fec_pick:
                fecha_seleccionada = st.date_input(
                    "Seleccione la fecha:",
                    value=date(2026, 4, 24)
                )
            with col_fec_txt:
                texto_fecha_auto = f"Maracaibo, {fecha_seleccionada.day} de {MESES_ES[fecha_seleccionada.month]} de {fecha_seleccionada.year}"
                lugar_fecha = st.text_input("Lugar y Fecha en el Certificado:", value=texto_fecha_auto)
            
            # Construcción dinámica del Código
            st.markdown("**Código de Registro / Control:**")
            col_pref, col_corr = st.columns([2, 1])
            
            with col_pref:
                lista_prefijos = [f"{c['prefijo']} - {c.get('descripcion', '')}" for c in codigos_guardados]
                if not lista_prefijos:
                    lista_prefijos = ["COQ-LCP-NPB - Nudo de Perseverancia en Bronce"]
                
                # Búsqueda inteligente del prefijo que coincida con la condecoración elegida
                default_pref_idx = 0
                if condec_obj:
                    nom_condec_words = set(w for w in condec_obj["nombre"].lower().replace("de ", "").replace("en ", "").split() if len(w) > 2)
                    for idx_c, c_item in enumerate(codigos_guardados):
                        desc_words = set(w for w in c_item.get("descripcion", "").lower().replace("de ", "").replace("en ", "").split() if len(w) > 2)
                        if nom_condec_words and len(nom_condec_words.intersection(desc_words)) >= 2:
                            default_pref_idx = idx_c
                            break
                
                pref_sel = st.selectbox(
                    "Prefijo Oficial (Distrito-Grupo-Siglas):",
                    lista_prefijos,
                    index=default_pref_idx
                )
                prefijo_puro = pref_sel.split(" - ")[0].strip()
                
            with col_corr:
                correlativo = st.number_input("N° Correlativo:", min_value=1, max_value=999, value=1, step=1)
                
            # Ensamblado automático: PREFIJO-AÑO-MES-CORRELATIVO
            ano_str = str(fecha_seleccionada.year)
            mes_str = f"{fecha_seleccionada.month:02d}"
            num_str = f"{correlativo:03d}"
            codigo_armado = f"{prefijo_puro}-{ano_str}-{mes_str}-{num_str}"
            
            codigo_control_final = st.text_input("Código de Control Final (editable):", value=codigo_armado)
            st.caption(f"Estructura oficial: `{prefijo_puro}` (3 partes) + `{ano_str}` (Año) + `{mes_str}` (Mes) + `{num_str}` (Correlativo)")
                
            st.markdown("---")
            st.subheader("3. Firmantes Oficiales")
            
            nombres_firmantes = ["(Ninguno)"] + [f["nombre"] for f in firmantes]
            col_izq, col_der = st.columns(2)
            with col_izq:
                firmante_izq_nombre = st.selectbox(
                    "Firmante Izquierdo:",
                    nombres_firmantes,
                    index=1 if len(nombres_firmantes) > 1 else 0
                )
            with col_der:
                firmante_der_nombre = st.selectbox(
                    "Firmante Derecho:",
                    nombres_firmantes,
                    index=2 if len(nombres_firmantes) > 2 else (1 if len(nombres_firmantes) > 1 else 0)
                )
                
            firmante_izq_obj = next((f for f in firmantes if f["nombre"] == firmante_izq_nombre), None)
            firmante_der_obj = next((f for f in firmantes if f["nombre"] == firmante_der_nombre), None)
            
            btn_generar = st.button("⚜️ Generar Certificado Oficial", type="primary", use_container_width=True)
            
        with col_prev:
            st.subheader("Vista Previa y Descarga")
            
            if btn_generar or "ultimo_pdf" not in st.session_state:
                if condec_obj and nombre_galardonado.strip():
                    ruta_template = os.path.join(PLANTILLAS_DIR, condec_obj["archivo_pdf"])
                    if os.path.exists(ruta_template):
                        with st.spinner("Procesando documento oficial..."):
                            pdf_bytes = generar_certificado_pdf(
                                template_path=ruta_template,
                                nombre=nombre_galardonado,
                                cedula=cedula,
                                lugar_fecha=lugar_fecha,
                                codigo_control=codigo_control_final,
                                firmante_izq=firmante_izq_obj,
                                firmante_der=firmante_der_obj,
                                firmas_dir=FIRMAS_DIR,
                                config=condec_obj.get("config")
                            )
                            st.session_state["ultimo_pdf"] = pdf_bytes
                            st.session_state["ultimo_nombre"] = nombre_galardonado.strip().replace(" ", "_")
                    else:
                        st.error(f"El archivo de plantilla '{condec_obj['archivo_pdf']}' no existe en el disco.")
            
            if "ultimo_pdf" in st.session_state:
                img_preview = pdf_to_image_bytes(st.session_state["ultimo_pdf"])
                st.image(img_preview, caption="Vista previa del certificado generado", use_container_width=True)
                
                nombre_archivo_descarga = f"Certificado_{st.session_state.get('ultimo_nombre', 'Condecoracion')}.pdf"
                st.download_button(
                    label="📥 Descargar Certificado Oficial en PDF",
                    data=st.session_state["ultimo_pdf"],
                    file_name=nombre_archivo_descarga,
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )

# -------------------------------------------------------------------------
# TAB 2: CÓDIGOS DE CONTROL
# -------------------------------------------------------------------------
with tab2:
    st.subheader("Configuración de Prefijos para Códigos de Control")
    st.markdown("""
    Los códigos oficiales se arman con la siguiente estructura:
    - **Parte 1 (Distrito):** ej. `COQ`
    - **Parte 2 (Grupo Scout):** ej. `LCP`
    - **Parte 3 (Siglas de Condecoración):** ej. `NPB` (Nudo de Perseverancia en Bronce)
    - **Parte 4 y 5 (Año y Mes):** Se toman automáticamente de la fecha del certificado (ej. `2026-04`)
    - **Parte 6 (Correlativo):** El número de condecoración entregado ese mes colocado a mano (ej. `001`)
    """)
    st.info("Aquí defines los prefijos de las 3 primeras partes para tenerlos listos en el formulario.")
    
    codigos_actuales = cargar_codigos()
    
    if codigos_actuales:
        st.markdown("### Prefijos Registrados")
        for idx, c in enumerate(codigos_actuales):
            with st.container():
                col_c1, col_c2, col_c3 = st.columns([3, 2, 1])
                with col_c1:
                    st.markdown(f"**Prefijo:** `{c['prefijo']}`")
                    st.caption(f"Descripción: {c.get('descripcion', 'Sin descripción')}")
                with col_c2:
                    st.markdown(f"Distrito: **{c.get('distrito','')}** | Grupo: **{c.get('grupo','')}** | Siglas: **{c.get('siglas','')}**")
                with col_c3:
                    if st.button("🗑️ Eliminar", key=f"del_cod_{idx}"):
                        codigos_actuales.pop(idx)
                        guardar_codigos(codigos_actuales)
                        st.success("Prefijo eliminado.")
                        st.rerun()
                st.divider()
    else:
        st.write("No hay prefijos registrados.")
        
    st.subheader("➕ Registrar Nuevo Prefijo de Código")
    with st.form("form_nuevo_codigo", clear_on_submit=True):
        col_d, col_g, col_s = st.columns(3)
        with col_d:
            nuevo_distrito = st.text_input("1. Distrito (ej. COQ):", max_chars=10).upper()
        with col_g:
            nuevo_grupo = st.text_input("2. Grupo Scout (ej. LCP):", max_chars=10).upper()
        with col_s:
            nuevas_siglas = st.text_input("3. Siglas Condecoración (ej. NPB):", max_chars=10).upper()
            
        nueva_desc = st.text_input("Descripción / Nombre de la Condecoración:", placeholder="ej. Nudo de Perseverancia en Bronce")
        
        btn_guardar_codigo = st.form_submit_button("Guardar Prefijo", type="primary")
        if btn_guardar_codigo:
            if not nuevo_distrito.strip() or not nuevo_grupo.strip() or not nuevas_siglas.strip():
                st.error("Por favor completa los 3 campos del código (Distrito, Grupo y Siglas).")
            else:
                prefijo_generado = f"{nuevo_distrito.strip()}-{nuevo_grupo.strip()}-{nuevas_siglas.strip()}"
                codigos_actuales.append({
                    "id": prefijo_generado.lower().replace("-", "_"),
                    "distrito": nuevo_distrito.strip(),
                    "grupo": nuevo_grupo.strip(),
                    "siglas": nuevas_siglas.strip(),
                    "prefijo": prefijo_generado,
                    "descripcion": nueva_desc.strip() or prefijo_generado
                })
                guardar_codigos(codigos_actuales)
                st.success(f"¡Prefijo '{prefijo_generado}' registrado exitosamente!")
                st.rerun()

# -------------------------------------------------------------------------
# TAB 3: GESTIÓN DE FIRMANTES
# -------------------------------------------------------------------------
with tab3:
    st.subheader("Autoridades y Firmantes Registrados")
    st.info("Aquí puedes cargar a los dirigentes y autoridades que firman las condecoraciones con sus cargos y firmas digitalizadas.")
    
    firmantes_actuales = cargar_firmantes()
    
    if firmantes_actuales:
        for idx, f in enumerate(firmantes_actuales):
            with st.container():
                col_info, col_img, col_del = st.columns([3, 2, 1])
                with col_info:
                    st.markdown(f"**Nombre:** {f['nombre']}")
                    st.markdown(f"**Cargo:** {f['cargo']}")
                    st.markdown(f"**Institución:** {f.get('entidad', 'Asociación de Scouts de Venezuela')}")
                with col_img:
                    ruta_img = os.path.join(FIRMAS_DIR, f.get("archivo_firma", ""))
                    if os.path.exists(ruta_img):
                        st.image(ruta_img, width=150, caption="Firma Digitalizada")
                    else:
                        st.caption("Sin imagen de firma")
                with col_del:
                    if st.button("🗑️ Eliminar", key=f"del_firm_{idx}"):
                        firmantes_actuales.pop(idx)
                        guardar_firmantes(firmantes_actuales)
                        st.success("Firmante eliminado.")
                        st.rerun()
                st.divider()
    else:
        st.write("No hay firmantes registrados.")
        
    st.subheader("➕ Agregar Nuevo Firmante")
    with st.form("form_nuevo_firmante", clear_on_submit=True):
        nuevo_nombre = st.text_input("Nombre Completo:")
        nuevo_cargo = st.text_input("Cargo Oficial:", placeholder="ej. JEFE DE GRUPO, COMISIONADO DISTRITAL")
        nueva_entidad = st.text_input("Institución / Entidad:", value="Asociación de Scouts de Venezuela")
        archivo_firma = st.file_uploader("Imagen de la Firma (PNG con fondo transparente recomendado):", type=["png", "jpg", "jpeg"])
        
        btn_guardar_firmante = st.form_submit_button("Guardar Firmante", type="primary")
        if btn_guardar_firmante:
            if not nuevo_nombre.strip() or not nuevo_cargo.strip():
                st.error("Por favor ingresa al menos el nombre y el cargo.")
            else:
                nombre_archivo = ""
                if archivo_firma:
                    nombre_archivo = f"{nuevo_nombre.lower().replace(' ', '_')}_{archivo_firma.name}"
                    ruta_destino = os.path.join(FIRMAS_DIR, nombre_archivo)
                    with open(ruta_destino, "wb") as f_out:
                        f_out.write(archivo_firma.getbuffer())
                
                firmantes_actuales.append({
                    "id": nuevo_nombre.lower().replace(" ", "_"),
                    "nombre": nuevo_nombre.strip(),
                    "cargo": nuevo_cargo.strip(),
                    "entidad": nueva_entidad.strip(),
                    "archivo_firma": nombre_archivo
                })
                guardar_firmantes(firmantes_actuales)
                st.success(f"Firmante '{nuevo_nombre}' agregado correctamente.")
                st.rerun()

# -------------------------------------------------------------------------
# TAB 4: GESTIÓN DE FORMATOS Y PLANTILLAS
# -------------------------------------------------------------------------
with tab4:
    st.subheader("Catálogo de Condecoraciones y Formatos PDF")
    st.info("Sube aquí tus plantillas PDF limpias. Si algún formato tiene una distribución diferente, puedes calibrar sus coordenadas individuales.")
    
    condec_actuales = cargar_condecoraciones()
    
    if condec_actuales:
        for idx, c in enumerate(condec_actuales):
            cfg = c.get("config", DEFAULT_CONFIG.copy())
            with st.container():
                col_cinfo, col_cprev, col_cdel = st.columns([3, 2, 1])
                with col_cinfo:
                    st.markdown(f"### {c['nombre']}")
                    st.caption(f"Archivo: `{c['archivo_pdf']}`")
                with col_cprev:
                    ruta_pdf = os.path.join(PLANTILLAS_DIR, c["archivo_pdf"])
                    if os.path.exists(ruta_pdf):
                        with open(ruta_pdf, "rb") as f_pdf:
                            pdf_img = pdf_to_image_bytes(f_pdf.read())
                            st.image(pdf_img, width=110, caption="Miniatura Formato")
                with col_cdel:
                    if st.button("🗑️ Eliminar Formato", key=f"del_condec_{idx}"):
                        condec_actuales.pop(idx)
                        guardar_condecoraciones(condec_actuales)
                        st.success("Formato eliminado.")
                        st.rerun()
                
                with st.expander(f"⚙️ Calibrar Posiciones para '{c['nombre']}'"):
                    st.caption("Ajusta las coordenadas si este formato tiene los textos o firmas en otra posición.")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown("**Nombre y Cédula**")
                        nuevo_nom_y = st.number_input(f"Altura Y del Nombre:", value=float(cfg.get("nombre_y", 467.0)), step=1.0, key=f"ny_{idx}")
                        nuevo_nom_tam = st.number_input(f"Tamaño máx. Nombre:", value=int(cfg.get("nombre_tamano_max", 36)), step=1, key=f"nt_{idx}")
                        nuevo_ced_x = st.number_input(f"Posición X Cédula:", value=float(cfg.get("cedula_x", 257.0)), step=1.0, key=f"cx_{idx}")
                        nuevo_ced_y = st.number_input(f"Altura Y Cédula:", value=float(cfg.get("cedula_y", 435.1)), step=0.5, key=f"cy_{idx}")
                        nuevo_ced_tam = st.number_input(f"Tamaño Cédula:", value=float(cfg.get("cedula_tamano", 16.5)), step=0.5, key=f"ct_{idx}")
                    with col2:
                        st.markdown("**Firmantes**")
                        nuevo_fizq_x = st.number_input(f"Centro X Firmante Izq:", value=float(cfg.get("firmante_izq_x", 160.0)), step=2.0, key=f"fix_{idx}")
                        nuevo_fder_x = st.number_input(f"Centro X Firmante Der:", value=float(cfg.get("firmante_der_x", 452.0)), step=2.0, key=f"fdx_{idx}")
                        nuevo_f_y = st.number_input(f"Altura Y Firmas:", value=float(cfg.get("firmante_izq_y", 138.0)), step=2.0, key=f"fy_{idx}")
                        nuevo_f_ancho = st.number_input(f"Ancho Firma:", value=float(cfg.get("ancho_firma", 135.0)), step=5.0, key=f"fa_{idx}")
                    with col3:
                        st.markdown("**Fecha y Código de Control**")
                        nuevo_fec_y = st.number_input(f"Altura Y Fecha:", value=float(cfg.get("lugar_fecha_y", 90.0)), step=1.0, key=f"ly_{idx}")
                        nuevo_cod_x = st.number_input(f"Centro X Código:", value=float(cfg.get("codigo_control_x", 178.2)), step=1.0, key=f"ccx_{idx}")
                        nuevo_cod_y = st.number_input(f"Altura Y Código:", value=float(cfg.get("codigo_control_y", 65.5)), step=0.5, key=f"ccy_{idx}")
                    
                    if st.button("💾 Guardar Nueva Calibración", key=f"btn_save_cfg_{idx}", type="primary"):
                        c["config"] = {
                            "nombre_y": nuevo_nom_y,
                            "nombre_tamano_max": nuevo_nom_tam,
                            "cedula_x": nuevo_ced_x,
                            "cedula_y": nuevo_ced_y,
                            "cedula_tamano": nuevo_ced_tam,
                            "firmante_izq_x": nuevo_fizq_x,
                            "firmante_izq_y": nuevo_f_y,
                            "firmante_der_x": nuevo_fder_x,
                            "firmante_der_y": nuevo_f_y,
                            "ancho_firma": nuevo_f_ancho,
                            "lugar_fecha_y": nuevo_fec_y,
                            "codigo_control_x": nuevo_cod_x,
                            "codigo_control_y": nuevo_cod_y,
                            "codigo_control_tamano": cfg.get("codigo_control_tamano", 7.5)
                        }
                        guardar_condecoraciones(condec_actuales)
                        st.success(f"¡Configuración de '{c['nombre']}' guardada correctamente!")
                        st.rerun()
                st.divider()
    else:
        st.write("No hay formatos registrados.")
        
    st.subheader("➕ Agregar Nuevo Formato Oficial")
    with st.form("form_nueva_condec", clear_on_submit=True):
        nombre_condec = st.text_input("Nombre de la Distinción / Condecoración:", placeholder="ej. Barra Espíritu Scout en Oro")
        pdf_subido = st.file_uploader("Archivo PDF oficial (Plantilla limpia):", type=["pdf"])
        
        btn_guardar_condec = st.form_submit_button("Guardar Formato", type="primary")
        if btn_guardar_condec:
            if not nombre_condec.strip() or not pdf_subido:
                st.error("Por favor ingresa el nombre y sube el archivo PDF.")
            else:
                nombre_archivo_pdf = f"{nombre_condec.lower().replace(' ', '_')}.pdf"
                ruta_destino_pdf = os.path.join(PLANTILLAS_DIR, nombre_archivo_pdf)
                with open(ruta_destino_pdf, "wb") as f_out:
                    f_out.write(pdf_subido.getbuffer())
                    
                condec_actuales.append({
                    "id": nombre_condec.lower().replace(" ", "_"),
                    "nombre": nombre_condec.strip(),
                    "archivo_pdf": nombre_archivo_pdf,
                    "config": DEFAULT_CONFIG.copy()
                })
                guardar_condecoraciones(condec_actuales)
                st.success(f"Condecoración '{nombre_condec}' agregada correctamente. Viene precalibrada con las posiciones estándar.")
                st.rerun()

# Pie de página Scout
st.markdown("""
<div class="scout-footer">
    ⚜️ <strong>Asociación de Scouts de Venezuela</strong> • Fundada en 1913 • Reconocida por WOSM<br>
    <em>"Construir un mundo mejor"</em> • Sistema Oficial de Condecoraciones
</div>
""", unsafe_allow_html=True)
