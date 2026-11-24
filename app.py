import os
import json
import io
import csv
import base64
import importlib
import streamlit as st
import openpyxl
from datetime import datetime, date
import generator
importlib.reload(generator)
from generator import (
    generar_certificado_pdf,
    pdf_to_image_bytes,
    unir_certificados_pdf,
    crear_zip_certificados,
    DEFAULT_CONFIG
)

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

# Configuración de página
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
REGISTRO_FILE = os.path.join(DATOS_DIR, "registro.json")

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

def cargar_registro():
    if os.path.exists(REGISTRO_FILE):
        with open(REGISTRO_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def guardar_registro(reg):
    with open(REGISTRO_FILE, "w", encoding="utf-8") as f:
        json.dump(reg, f, indent=2, ensure_ascii=False)

def registrar_emision(nombre, cedula, condecoracion, codigo, lugar_fecha, firmante_izq, firmante_der):
    reg = cargar_registro()
    nuevo_item = {
        "id": str(len(reg) + 1),
        "fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "codigo": codigo.strip(),
        "condecoracion": condecoracion.strip(),
        "nombre": nombre.strip().upper(),
        "cedula": cedula.strip(),
        "lugar_fecha": lugar_fecha.strip(),
        "firmante_izq": firmante_izq.get("nombre", "") if firmante_izq else "",
        "cargo_izq": firmante_izq.get("cargo", "") if firmante_izq else "",
        "firmante_der": firmante_der.get("nombre", "") if firmante_der else "",
        "cargo_der": firmante_der.get("cargo", "") if firmante_der else ""
    }
    if not any(r.get("codigo") == nuevo_item["codigo"] and r.get("cedula") == nuevo_item["cedula"] for r in reg):
        reg.insert(0, nuevo_item)
        guardar_registro(reg)

def normalizar_columna(col_nombre: str) -> str:
    c = str(col_nombre).strip().lower().replace("_", " ")
    if any(k in c for k in ["nombre", "galardonado", "beneficiario", "scout"]):
        return "Nombre"
    if any(k in c for k in ["cedula", "cédula", "ci", "c.i.", "documento", "identidad"]):
        return "Cedula"
    if any(k in c for k in ["condecoracion", "condecoración", "distincion", "distinción", "tipo", "premio"]):
        return "Condecoracion"
    if any(k in c for k in ["correlativo", "numero", "número", "nro", "num"]):
        return "Correlativo"
    return str(col_nombre).strip().title()

def leer_archivo_lote(uploaded_file):
    """
    Lee archivos Excel (.xlsx/.xls) o CSV usando librerías ligeras sin requerir pandas.
    """
    filas = []
    fname = uploaded_file.name.lower()
    
    if fname.endswith(".csv"):
        contenido = uploaded_file.getvalue().decode("utf-8", errors="ignore")
        reader = csv.reader(io.StringIO(contenido))
        header = next(reader, None)
        if header:
            norm_header = [normalizar_columna(h) for h in header]
            for r in reader:
                if any(cell.strip() for cell in r):
                    fila_dict = {}
                    for i, h in enumerate(norm_header):
                        fila_dict[h] = r[i].strip() if i < len(r) else ""
                    filas.append(fila_dict)
    else:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        sheet = wb.active
        raw_rows = list(sheet.iter_rows(values_only=True))
        if raw_rows:
            header = raw_rows[0]
            norm_header = [normalizar_columna(h) for h in header if h is not None]
            for r in raw_rows[1:]:
                if any(cell is not None and str(cell).strip() for cell in r):
                    fila_dict = {}
                    for i, h in enumerate(norm_header):
                        val = r[i] if i < len(r) else ""
                        fila_dict[h] = str(val).strip() if val is not None else ""
                    filas.append(fila_dict)
    return filas

def crear_excel_ejemplo():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Galardonados"
    ws.append(["Nombre", "Cedula", "Condecoracion", "Correlativo"])
    ws.append(["CARLOS ALBERTO ESPINOZA", "33.792.612", "Barra Espíritu Scout en Bronce", 1])
    ws.append(["MARIA ELENA CHACON", "28.450.119", "Nudo de Servicio en Oro", 2])
    ws.append(["JOSE LUIS URDANETA", "31.902.833", "Barra de Liderazgo", 3])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def exportar_registro_excel(registros):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Libro_Oficial_ASV"
    headers = [
        "N° Registro", "Fecha y Hora", "Código de Control", "Condecoración",
        "Galardonado", "Cédula", "Lugar y Fecha",
        "Firmante Izquierdo", "Cargo Izq", "Firmante Derecho", "Cargo Der"
    ]
    ws.append(headers)
    for r in registros:
        ws.append([
            r.get("id", ""),
            r.get("fecha_registro", ""),
            r.get("codigo", ""),
            r.get("condecoracion", ""),
            r.get("nombre", ""),
            r.get("cedula", ""),
            r.get("lugar_fecha", ""),
            r.get("firmante_izq", ""),
            r.get("cargo_izq", ""),
            r.get("firmante_der", ""),
            r.get("cargo_der", "")
        ])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# Estilos personalizados Scouts
logo_b64 = get_base64_image(LOGO_SCOUT_PATH)
st.markdown(f"""
<style>
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

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📜 Emitir Certificado",
    "📖 Libro de Registro (Art. 35)",
    "🏷️ Códigos de Control",
    "✍️ Gestión de Firmantes",
    "📁 Formatos y Plantillas"
])

# -------------------------------------------------------------------------
# TAB 1: EMITIR CERTIFICADO (INDIVIDUAL Y POR LOTE)
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
        modo_emision = st.radio(
            "Modalidad de Emisión:",
            ["Individual (Uno a Uno)", "⚡ Emisión Masiva por Lote (Desde Excel)"],
            horizontal=True
        )
        st.markdown("---")
        
        # MODO INDIVIDUAL
        if modo_emision == "Individual (Uno a Uno)":
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
                    fecha_seleccionada = st.date_input("Seleccione la fecha:", value=date(2026, 4, 24))
                with col_fec_txt:
                    texto_fecha_auto = f"Maracaibo, {fecha_seleccionada.day} de {MESES_ES[fecha_seleccionada.month]} de {fecha_seleccionada.year}"
                    lugar_fecha = st.text_input("Lugar y Fecha en el Certificado:", value=texto_fecha_auto)
                
                st.markdown("**Código de Registro / Control:**")
                col_pref, col_corr = st.columns([2, 1])
                
                with col_pref:
                    lista_prefijos = [f"{c['prefijo']} - {c.get('descripcion', '')}" for c in codigos_guardados]
                    if not lista_prefijos:
                        lista_prefijos = ["COQ-LCP-NPB - Nudo de Perseverancia en Bronce"]
                    
                    default_pref_idx = 0
                    if condec_obj:
                        nom_words = set(w for w in condec_obj["nombre"].lower().replace("de ", "").replace("en ", "").split() if len(w) > 2)
                        for idx_c, c_item in enumerate(codigos_guardados):
                            desc_words = set(w for w in c_item.get("descripcion", "").lower().replace("de ", "").replace("en ", "").split() if len(w) > 2)
                            if nom_words and len(nom_words.intersection(desc_words)) >= 2:
                                default_pref_idx = idx_c
                                break
                    
                    pref_sel = st.selectbox("Prefijo Oficial (Distrito-Grupo-Siglas):", lista_prefijos, index=default_pref_idx)
                    prefijo_puro = pref_sel.split(" - ")[0].strip()
                    
                with col_corr:
                    correlativo = st.number_input("N° Correlativo:", min_value=1, max_value=999, value=1, step=1)
                    
                ano_str = str(fecha_seleccionada.year)
                mes_str = f"{fecha_seleccionada.month:02d}"
                num_str = f"{correlativo:03d}"
                codigo_armado = f"{prefijo_puro}-{ano_str}-{mes_str}-{num_str}"
                
                codigo_control_final = st.text_input("Código de Control Final (editable):", value=codigo_armado)
                st.caption(f"Estructura: `{prefijo_puro}` + `{ano_str}` + `{mes_str}` + `{num_str}`")
                    
                st.markdown("---")
                st.subheader("3. Firmantes Oficiales")
                
                nombres_firmantes = ["(Ninguno)"] + [f["nombre"] for f in firmantes]
                col_izq, col_der = st.columns(2)
                with col_izq:
                    firmante_izq_nombre = st.selectbox("Firmante Izquierdo:", nombres_firmantes, index=1 if len(nombres_firmantes) > 1 else 0)
                with col_der:
                    firmante_der_nombre = st.selectbox("Firmante Derecho:", nombres_firmantes, index=2 if len(nombres_firmantes) > 2 else (1 if len(nombres_firmantes) > 1 else 0))
                    
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
                                
                                registrar_emision(
                                    nombre=nombre_galardonado,
                                    cedula=cedula,
                                    condecoracion=condec_obj["nombre"],
                                    codigo=codigo_control_final,
                                    lugar_fecha=lugar_fecha,
                                    firmante_izq=firmante_izq_obj,
                                    firmante_der=firmante_der_obj
                                )
                        else:
                            st.error(f"El archivo de plantilla '{condec_obj['archivo_pdf']}' no existe.")
                
                if "ultimo_pdf" in st.session_state:
                    img_preview = pdf_to_image_bytes(st.session_state["ultimo_pdf"])
                    st.image(img_preview, caption="Vista previa del certificado generado", use_container_width=True)
                    
                    nombre_descarga = f"Certificado_{st.session_state.get('ultimo_nombre', 'Condecoracion')}.pdf"
                    st.download_button(
                        label="📥 Descargar Certificado Oficial en PDF",
                        data=st.session_state["ultimo_pdf"],
                        file_name=nombre_descarga,
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )

        # MODO MASIVO POR LOTE
        else:
            st.subheader("⚡ Emisión Masiva de Certificados por Lote")
            st.markdown("""
            Esta herramienta te permite emitir decenas de condecoraciones a la vez a partir de una lista en Excel o CSV.
            El Excel puede llevar: **Nombre**, **Cédula**, y opcionalmente **Condecoración** y **Correlativo**.
            """)
            
            # Plantilla descargable de Excel
            buf_ejemplo = crear_excel_ejemplo()
            st.download_button(
                label="📄 Descargar Plantilla Excel de Ejemplo (.xlsx)",
                data=buf_ejemplo,
                file_name="plantilla_condecoraciones_scouts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("---")
            col_lote_cfg, col_lote_file = st.columns([1, 1], gap="large")
            
            with col_lote_cfg:
                st.markdown("#### Configuración Predeterminada del Lote")
                nombres_condec = [c["nombre"] for c in condecoraciones]
                condec_lote_sel = st.selectbox(
                    "Condecoración (para filas sin condecoración explícita):",
                    nombres_condec,
                    key="condec_lote"
                )
                condec_lote_obj = next((c for c in condecoraciones if c["nombre"] == condec_lote_sel), None)
                
                col_lf_p, col_lf_t = st.columns(2)
                with col_lf_p:
                    fec_lote = st.date_input("Fecha de emisión:", value=date(2026, 4, 24), key="fec_lote")
                with col_lf_t:
                    txt_fec_lote = f"Maracaibo, {fec_lote.day} de {MESES_ES[fec_lote.month]} de {fec_lote.year}"
                    lugar_fecha_lote = st.text_input("Lugar y Fecha:", value=txt_fec_lote, key="lf_lote_txt")
                
                lista_prefijos = [f"{c['prefijo']} - {c.get('descripcion', '')}" for c in codigos_guardados]
                pref_lote_sel = st.selectbox("Prefijo Oficial para el Lote:", lista_prefijos, key="pref_lote")
                prefijo_lote_puro = pref_lote_sel.split(" - ")[0].strip()
                
                st.markdown("#### Firmantes Oficiales para el Lote")
                nombres_firmantes = ["(Ninguno)"] + [f["nombre"] for f in firmantes]
                col_fl_izq, col_fl_der = st.columns(2)
                with col_fl_izq:
                    fl_izq_nombre = st.selectbox("Firmante Izq:", nombres_firmantes, index=1 if len(nombres_firmantes) > 1 else 0, key="fl_izq")
                with col_fl_der:
                    fl_der_nombre = st.selectbox("Firmante Der:", nombres_firmantes, index=2 if len(nombres_firmantes) > 2 else 0, key="fl_der")
                    
                firm_lote_izq = next((f for f in firmantes if f["nombre"] == fl_izq_nombre), None)
                firm_lote_der = next((f for f in firmantes if f["nombre"] == fl_der_nombre), None)
                
            with col_lote_file:
                st.markdown("#### Cargar Archivo de Galardonados")
                archivo_lote = st.file_uploader("Subir archivo Excel (.xlsx) o CSV:", type=["xlsx", "xls", "csv"])
                
                if archivo_lote:
                    try:
                        filas_cargadas = leer_archivo_lote(archivo_lote)
                        if not filas_cargadas:
                            st.error("No se encontraron filas con datos en el archivo subido.")
                        else:
                            st.success(f"Archivo cargado: **{len(filas_cargadas)} personas** detectadas.")
                            st.dataframe(filas_cargadas[:10], use_container_width=True)
                            
                            btn_procesar_lote = st.button("🚀 Procesar y Generar Lote Completo", type="primary", use_container_width=True)
                            
                            if btn_procesar_lote:
                                progreso = st.progress(0)
                                status_txt = st.empty()
                                
                                lista_pdfs_bytes = []
                                lista_para_zip = []
                                ano_str = str(fec_lote.year)
                                mes_str = f"{fec_lote.month:02d}"
                                total_filas = len(filas_cargadas)
                                
                                for i, row in enumerate(filas_cargadas):
                                    status_txt.text(f"Generando certificado {i+1} de {total_filas}...")
                                    nom_persona = str(row.get("Nombre", "")).strip().upper()
                                    ced_persona = str(row.get("Cedula", "")).strip()
                                    
                                    # Determinar condecoración para esta fila
                                    condec_fila_nombre = str(row.get("Condecoracion", "")).strip()
                                    target_condec = None
                                    if condec_fila_nombre:
                                        # Buscar coincidencia
                                        target_condec = next((c for c in condecoraciones if condec_fila_nombre.lower() in c["nombre"].lower()), None)
                                    if not target_condec:
                                        target_condec = condec_lote_obj
                                        
                                    ruta_template = os.path.join(PLANTILLAS_DIR, target_condec["archivo_pdf"])
                                    if not os.path.exists(ruta_template):
                                        st.warning(f"Plantilla '{target_condec['archivo_pdf']}' no encontrada para {nom_persona}. Saltando...")
                                        continue
                                        
                                    # Correlativo
                                    num_corr = row.get("Correlativo", i + 1)
                                    try:
                                        num_corr_int = int(num_corr)
                                    except:
                                        num_corr_int = i + 1
                                    cod_armado = f"{prefijo_lote_puro}-{ano_str}-{mes_str}-{num_corr_int:03d}"
                                    
                                    pdf_ind = generar_certificado_pdf(
                                        template_path=ruta_template,
                                        nombre=nom_persona,
                                        cedula=ced_persona,
                                        lugar_fecha=lugar_fecha_lote,
                                        codigo_control=cod_armado,
                                        firmante_izq=firm_lote_izq,
                                        firmante_der=firm_lote_der,
                                        firmas_dir=FIRMAS_DIR,
                                        config=target_condec.get("config")
                                    )
                                    lista_pdfs_bytes.append(pdf_ind)
                                    nombre_archivo_ind = f"Certificado_{num_corr_int:03d}_{nom_persona.replace(' ', '_')}.pdf"
                                    lista_para_zip.append((nombre_archivo_ind, pdf_ind))
                                    
                                    registrar_emision(
                                        nombre=nom_persona,
                                        cedula=ced_persona,
                                        condecoracion=target_condec["nombre"],
                                        codigo=cod_armado,
                                        lugar_fecha=lugar_fecha_lote,
                                        firmante_izq=firm_lote_izq,
                                        firmante_der=firm_lote_der
                                    )
                                    progreso.progress((i + 1) / total_filas)
                                    
                                if lista_pdfs_bytes:
                                    status_txt.success(f"¡{len(lista_pdfs_bytes)} certificados generados con éxito!")
                                    pdf_unificado = unir_certificados_pdf(lista_pdfs_bytes)
                                    zip_individuales = crear_zip_certificados(lista_para_zip)
                                    
                                    st.markdown("---")
                                    st.markdown("### 📥 Descargas del Lote")
                                    col_d1, col_d2 = st.columns(2)
                                    with col_d1:
                                        st.download_button(
                                            label=f"📄 Descargar Todo en 1 Solo PDF ({len(lista_pdfs_bytes)} páginas)",
                                            data=pdf_unificado,
                                            file_name=f"Lote_Condecoraciones_{fec_lote.strftime('%Y%m%d')}.pdf",
                                            mime="application/pdf",
                                            type="primary",
                                            use_container_width=True
                                        )
                                    with col_d2:
                                        st.download_button(
                                            label=f"📦 Descargar Archivo ZIP (PDFs individuales)",
                                            data=zip_individuales,
                                            file_name=f"Lote_Condecoraciones_{fec_lote.strftime('%Y%m%d')}.zip",
                                            mime="application/zip",
                                            use_container_width=True
                                        )
                    except Exception as e:
                        st.error(f"Error al procesar el archivo: {e}")

# -------------------------------------------------------------------------
# TAB 2: LIBRO DE REGISTRO (ART. 35)
# -------------------------------------------------------------------------
with tab2:
    st.subheader("📖 Libro de Registro Oficial de Condecoraciones y Distinciones")
    st.markdown("""
    > *Cumplimiento del **Artículo 35 del Reglamento Oficial**: Se hace constar el registro oficial de cada condecoración otorgada, antecedentes, fecha y número de certificado.*
    """)
    
    registro_actual = cargar_registro()
    
    if registro_actual:
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Condecoraciones Emitidas", len(registro_actual))
        m2.metric("Última Emisión", registro_actual[0]["fecha_registro"][:10] if registro_actual else "-")
        condec_unicas = len(set(r.get("condecoracion", "") for r in registro_actual))
        m3.metric("Condecoraciones Diferentes", condec_unicas)
        
        st.markdown("---")
        
        col_filtro, col_exp = st.columns([3, 1])
        with col_filtro:
            texto_busqueda = st.text_input("🔍 Buscar por Nombre, Cédula o Código de Registro:", placeholder="ej. Silverio, 11.289.602, NPB...")
        with col_exp:
            buf_excel = exportar_registro_excel(registro_actual)
            st.download_button(
                label="📥 Exportar a Excel (.xlsx)",
                data=buf_excel,
                file_name=f"Libro_Registro_Condecoraciones_ASV_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
        if texto_busqueda.strip():
            filtro = texto_busqueda.strip().lower()
            mostrar = [
                r for r in registro_actual
                if filtro in r.get("nombre", "").lower()
                or filtro in r.get("cedula", "").lower()
                or filtro in r.get("codigo", "").lower()
                or filtro in r.get("condecoracion", "").lower()
            ]
        else:
            mostrar = registro_actual
            
        st.dataframe(mostrar, use_container_width=True, hide_index=True)
        
        with st.expander("🗑️ Opciones de Mantenimiento"):
            if st.button("Limpiar Historial de Pruebas (Borrar todo el registro)"):
                guardar_registro([])
                st.success("Historial reiniciado.")
                st.rerun()
    else:
        st.info("Aún no se han emitido certificados. Cuando generes el primero (individual o por lote), aparecerá registrado aquí automáticamente.")

# -------------------------------------------------------------------------
# TAB 3: CÓDIGOS DE CONTROL
# -------------------------------------------------------------------------
with tab3:
    st.subheader("Configuración de Prefijos para Códigos de Control")
    st.markdown("""
    Los códigos oficiales se arman con la siguiente estructura:
    - **Parte 1 (Distrito):** ej. `COQ`
    - **Parte 2 (Grupo Scout):** ej. `LCP`
    - **Parte 3 (Siglas de Condecoración):** ej. `NPB` (Nudo de Perseverancia en Bronce)
    - **Parte 4 y 5 (Año y Mes):** Se toman automáticamente de la fecha del certificado (ej. `2026-04`)
    - **Parte 6 (Correlativo):** El número de condecoración entregado ese mes colocado a mano o automático (ej. `001`)
    """)
    st.info("Aquí defines los prefijos de las 3 primeras partes para tenerlos listos en el formulario.")
    
    codigos_actuales = cargar_codigos()
    
    if codigos_actuales:
        st.markdown(f"### Prefijos Registrados ({len(codigos_actuales)} disponibles)")
        
        busq_cod = st.text_input("Filtrar prefijos por nombre o sigla:", placeholder="ej. Nudo, Barra, Oro, Bronce...")
        lista_mostrar = codigos_actuales
        if busq_cod.strip():
            b = busq_cod.strip().lower()
            lista_mostrar = [c for c in codigos_actuales if b in c["prefijo"].lower() or b in c.get("descripcion", "").lower()]
            
        for idx, c in enumerate(lista_mostrar):
            with st.container():
                col_c1, col_c2, col_c3 = st.columns([3, 2, 1])
                with col_c1:
                    st.markdown(f"**Prefijo:** `{c['prefijo']}`")
                    st.caption(f"Descripción: {c.get('descripcion', 'Sin descripción')}")
                with col_c2:
                    st.markdown(f"Distrito: **{c.get('distrito','')}** | Grupo: **{c.get('grupo','')}** | Siglas: **{c.get('siglas','')}**")
                with col_c3:
                    if st.button("🗑️ Eliminar", key=f"del_cod_{idx}"):
                        codigos_actuales.remove(c)
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
# TAB 4: GESTIÓN DE FIRMANTES
# -------------------------------------------------------------------------
with tab4:
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
# TAB 5: GESTIÓN DE FORMATOS Y PLANTILLAS (INDIVIDUAL Y POR LOTE)
# -------------------------------------------------------------------------
with tab5:
    st.subheader("Catálogo de Condecoraciones y Formatos PDF")
    st.info("Aquí gestionas las plantillas oficiales. Puedes subir archivos individuales o múltiples plantillas a la vez y calibrarlas por lote.")
    
    condec_actuales = cargar_condecoraciones()
    
    if len(condec_actuales) > 1:
        with st.expander("⚙️ Calibración por Lote (Copiar coordenadas a múltiples formatos)"):
            st.markdown("""
            Si tienes varios formatos que comparten la misma estructura oficial, puedes tomar la calibración de uno y aplicarla a todos los demás de un solo clic.
            """)
            col_m1, col_m2 = st.columns([1, 2])
            with col_m1:
                formato_modelo = st.selectbox("Copiar calibración DESDE:", [c["nombre"] for c in condec_actuales], key="modelo_lote")
                formato_modelo_obj = next((c for c in condec_actuales if c["nombre"] == formato_modelo), None)
            with col_m2:
                formatos_destino = st.multiselect("Aplicar esta calibración HACIA los siguientes formatos:", [c["nombre"] for c in condec_actuales if c["nombre"] != formato_modelo], key="destinos_lote")
            
            if st.button("🚀 Aplicar Calibración a los Formatos Seleccionados", type="primary"):
                if formato_modelo_obj and formatos_destino:
                    cfg_copiar = formato_modelo_obj.get("config", DEFAULT_CONFIG).copy()
                    for c in condec_actuales:
                        if c["nombre"] in formatos_destino:
                            c["config"] = cfg_copiar.copy()
                    guardar_condecoraciones(condec_actuales)
                    st.success(f"¡Calibración de '{formato_modelo}' aplicada exitosamente a {len(formatos_destino)} formatos!")
                    st.rerun()
                else:
                    st.warning("Selecciona al menos un formato de destino.")
    
    st.markdown("---")
    
    if condec_actuales:
        st.markdown(f"### Formatos Oficiales Registrados ({len(condec_actuales)})")
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
        
    st.markdown("---")
    
    st.subheader("➕ Cargar Formatos Oficiales (Individual o por Lote)")
    tipo_subida = st.radio("Método de subida:", ["Subir un solo formato", "📦 Subir Múltiples Formatos a la vez (Lote)"], horizontal=True)
    
    if tipo_subida == "Subir un solo formato":
        with st.form("form_nueva_condec_ind", clear_on_submit=True):
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
                    st.success(f"Condecoración '{nombre_condec}' agregada correctamente.")
                    st.rerun()
    else:
        st.markdown("""
        Selecciona todos los archivos PDF que quieras cargar a la vez. El sistema usará el nombre del archivo para identificar la condecoración y le aplicará la calibración estándar automáticamente.
        """)
        pdfs_multiples = st.file_uploader(
            "Selecciona múltiples archivos PDF de plantillas oficiales:",
            type=["pdf"],
            accept_multiple_files=True
        )
        if pdfs_multiples:
            st.info(f"Se seleccionaron **{len(pdfs_multiples)} archivos** para subir.")
            if st.button("🚀 Subir y Registrar Todo el Lote de Formatos", type="primary"):
                agregados = 0
                for pdf_file in pdfs_multiples:
                    nombre_base = os.path.splitext(pdf_file.name)[0]
                    nombre_formateado = nombre_base.replace("_", " ").replace("-", " ").title()
                    nombre_archivo_pdf = f"{nombre_base.lower().replace(' ', '_')}.pdf"
                    
                    ruta_destino_pdf = os.path.join(PLANTILLAS_DIR, nombre_archivo_pdf)
                    with open(ruta_destino_pdf, "wb") as f_out:
                        f_out.write(pdf_file.getbuffer())
                        
                    existente = next((c for c in condec_actuales if c["archivo_pdf"] == nombre_archivo_pdf), None)
                    if not existente:
                        condec_actuales.append({
                            "id": nombre_base.lower().replace(" ", "_"),
                            "nombre": nombre_formateado,
                            "archivo_pdf": nombre_archivo_pdf,
                            "config": DEFAULT_CONFIG.copy()
                        })
                        agregados += 1
                        
                guardar_condecoraciones(condec_actuales)
                st.success(f"¡Se agregaron con éxito **{agregados} nuevos formatos** al catálogo!")
                st.rerun()

# Pie de página Scout
st.markdown("""
<div class="scout-footer">
    ⚜️ <strong>Asociación de Scouts de Venezuela</strong> • Fundada en 1913 • Reconocida por WOSM<br>
    <em>"Construir un mundo mejor"</em> • Sistema Oficial de Condecoraciones
</div>
""", unsafe_allow_html=True)
