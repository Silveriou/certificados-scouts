import os
import json
import importlib
import streamlit as st
from datetime import datetime
import generator
importlib.reload(generator)
from generator import generar_certificado_pdf, pdf_to_image_bytes

# Configuración de página
st.set_page_config(
    page_title="Sistema de Emisión de Condecoraciones",
    page_icon="⚜️",
    layout="wide"
)

# Rutas base
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PLANTILLAS_DIR = os.path.join(BASE_DIR, "plantillas")
FIRMAS_DIR = os.path.join(BASE_DIR, "firmas")
DATOS_DIR = os.path.join(BASE_DIR, "datos")
GENERADOS_DIR = os.path.join(BASE_DIR, "generados")

for d in [PLANTILLAS_DIR, FIRMAS_DIR, DATOS_DIR, GENERADOS_DIR]:
    os.makedirs(d, exist_ok=True)

FIRMANTES_FILE = os.path.join(DATOS_DIR, "firmantes.json")
CONDECORACIONES_FILE = os.path.join(DATOS_DIR, "condecoraciones.json")

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
            return json.load(f)
    return []

def guardar_condecoraciones(condec):
    with open(CONDECORACIONES_FILE, "w", encoding="utf-8") as f:
        json.dump(condec, f, indent=2, ensure_ascii=False)

# Encabezado
st.title("⚜️ Sistema de Emisión de Condecoraciones")
st.caption("Generación oficial de certificados y distinciones en alta fidelidad")

tab1, tab2, tab3 = st.tabs([
    "📜 Emitir Certificado",
    "✍️ Gestión de Firmantes",
    "📁 Formatos y Plantillas"
])

# -------------------------------------------------------------------------
# TAB 1: EMITIR CERTIFICADO
# -------------------------------------------------------------------------
with tab1:
    condecoraciones = cargar_condecoraciones()
    firmantes = cargar_firmantes()
    
    if not condecoraciones:
        st.warning("No hay formatos de condecoración cargados. Por favor ve a la pestaña 'Formatos y Plantillas' para agregar uno.")
    elif not firmantes:
        st.warning("No hay firmantes registrados. Por favor ve a la pestaña 'Gestión de Firmantes' para agregar al menos uno.")
    else:
        col_form, col_prev = st.columns([1, 1], gap="large")
        
        with col_form:
            st.subheader("Datos de la Condecoración")
            
            nombres_condec = [c["nombre"] for c in condecoraciones]
            condec_seleccionada = st.selectbox("Seleccione la Condecoración / Formato:", nombres_condec)
            condec_obj = next((c for c in condecoraciones if c["nombre"] == condec_seleccionada), None)
            
            st.markdown("---")
            st.subheader("Datos del Galardonado")
            
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
            
            col_fec, col_cod = st.columns(2)
            with col_fec:
                lugar_fecha = st.text_input(
                    "Lugar y Fecha:",
                    value="Maracaibo, 24 de abril de 2026"
                )
            with col_cod:
                codigo_control = st.text_input(
                    "Código de Registro / Control:",
                    value="COQ-LCP-NPB-2026-04-001"
                )
                
            st.markdown("---")
            st.subheader("Firmantes Oficiales")
            
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
            
            btn_generar = st.button("🚀 Generar Certificado", type="primary", use_container_width=True)
            
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
                                codigo_control=codigo_control,
                                firmante_izq=firmante_izq_obj,
                                firmante_der=firmante_der_obj,
                                firmas_dir=FIRMAS_DIR
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
# TAB 2: GESTIÓN DE FIRMANTES
# -------------------------------------------------------------------------
with tab2:
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
# TAB 3: GESTIÓN DE FORMATOS Y PLANTILLAS
# -------------------------------------------------------------------------
with tab3:
    st.subheader("Catálogo de Condecoraciones y Formatos PDF")
    st.info("Aquí puedes cargar los formatos oficiales (PDF limpios) de cada condecoración: Oro, Plata, Bronce, etc.")
    
    condec_actuales = cargar_condecoraciones()
    
    if condec_actuales:
        for idx, c in enumerate(condec_actuales):
            with st.container():
                col_cinfo, col_cprev, col_cdel = st.columns([3, 2, 1])
                with col_cinfo:
                    st.markdown(f"**Nombre:** {c['nombre']}")
                    st.caption(f"Archivo: {c['archivo_pdf']}")
                with col_cprev:
                    ruta_pdf = os.path.join(PLANTILLAS_DIR, c["archivo_pdf"])
                    if os.path.exists(ruta_pdf):
                        with open(ruta_pdf, "rb") as f_pdf:
                            pdf_img = pdf_to_image_bytes(f_pdf.read())
                            st.image(pdf_img, width=120, caption="Miniatura Formato")
                with col_cdel:
                    if st.button("🗑️ Eliminar", key=f"del_condec_{idx}"):
                        condec_actuales.pop(idx)
                        guardar_condecoraciones(condec_actuales)
                        st.success("Formato eliminado.")
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
                    "archivo_pdf": nombre_archivo_pdf
                })
                guardar_condecoraciones(condec_actuales)
                st.success(f"Condecoración '{nombre_condec}' agregada correctamente.")
                st.rerun()
