import base64
import csv
import hashlib
import hmac
import io
import json
import os
import re
import secrets
import shutil
import time
import unicodedata
from datetime import datetime, timedelta, timezone

import openpyxl
import streamlit as st

from generator import (
    DEFAULT_CONFIG,
    componer_paginas_carta,
    crear_respaldo_sistema_zip,
    crear_zip_certificados,
    formatear_cedula_vzla,
    generar_certificado_pdf,
    limpiar_pdf_lleno,
    limpiar_y_convertir_docx_a_pdf,
    pdf_to_image_bytes,
    restaurar_respaldo_sistema_zip,
)

# -------------------------------------------------------------------------
# RUTAS Y CONSTANTES
# -------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
PLANTILLAS_DIR = os.path.join(BASE_DIR, "plantillas")
ORIGINALES_DIR = os.path.join(PLANTILLAS_DIR, "_originales")
FIRMAS_DIR = os.path.join(BASE_DIR, "firmas")
DATOS_DIR = os.path.join(BASE_DIR, "datos")

for d in [ASSETS_DIR, PLANTILLAS_DIR, ORIGINALES_DIR, FIRMAS_DIR, DATOS_DIR]:
    os.makedirs(d, exist_ok=True)

LOGO_SCOUT_PATH = os.path.join(ASSETS_DIR, "logo_scout.png")
FIRMANTES_FILE = os.path.join(DATOS_DIR, "firmantes.json")
CONDECORACIONES_FILE = os.path.join(DATOS_DIR, "condecoraciones.json")
CODIGOS_FILE = os.path.join(DATOS_DIR, "codigos.json")
REGISTRO_FILE = os.path.join(DATOS_DIR, "registro.json")
USUARIOS_FILE = os.path.join(DATOS_DIR, "usuarios.json")

CIUDAD = "Maracaibo"
TZ_VE = timezone(timedelta(hours=-4))  # Venezuela: UTC-4 fijo, sin horario de verano
MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

ROLES = {
    "admin": {
        "nombre": "Administrador",
        "descripcion": "Control total: usuarios, respaldo y restauración, limpieza del historial.",
        "permisos": {"emitir", "registro", "anular", "limpiar", "codigos", "firmantes", "formatos", "respaldo", "usuarios"},
    },
    "gestor": {
        "nombre": "Gestor",
        "descripcion": "Emite, anula registros y gestiona códigos, firmantes y formatos.",
        "permisos": {"emitir", "registro", "anular", "codigos", "firmantes", "formatos"},
    },
    "emisor": {
        "nombre": "Emisor",
        "descripcion": "Emite certificados y consulta el Libro de Registro.",
        "permisos": {"emitir", "registro"},
    },
    "consulta": {
        "nombre": "Consulta",
        "descripcion": "Solo ve y exporta el Libro de Registro.",
        "permisos": {"registro"},
    },
}

st.set_page_config(
    page_title="ASV | Condecoraciones Scouts",
    page_icon=LOGO_SCOUT_PATH if os.path.exists(LOGO_SCOUT_PATH) else "⚜️",
    layout="wide",
)


# -------------------------------------------------------------------------
# UTILIDADES
# -------------------------------------------------------------------------
def ahora() -> datetime:
    return datetime.now(TZ_VE)


def sin_acentos(texto) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", str(texto)) if not unicodedata.combining(ch))


def normalizar(texto) -> str:
    return re.sub(r"\s+", " ", sin_acentos(texto).lower()).strip()


def slug(texto) -> str:
    s = re.sub(r"[^a-z0-9]+", "_", sin_acentos(texto).lower()).strip("_")
    return s or "sin_nombre"


def texto_lugar_fecha(fecha) -> str:
    return f"{CIUDAD}, {fecha.day} de {MESES_ES[fecha.month]} de {fecha.year}"


def leer_json(ruta, defecto):
    if os.path.exists(ruta):
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    return defecto


def escribir_json(ruta, datos):
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(datos, f, indent=2, ensure_ascii=False)
    os.replace(tmp, ruta)


def avisar(mensaje: str, tipo: str = "success"):
    """Mensaje que sobrevive al st.rerun() y se muestra arriba en la siguiente carga."""
    st.session_state.setdefault("_avisos", []).append((tipo, mensaje))


def mostrar_avisos():
    for tipo, mensaje in st.session_state.pop("_avisos", []):
        getattr(st, tipo)(mensaje)


def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""


@st.cache_data(show_spinner=False)
def miniatura_pdf(ruta: str, mtime: float) -> bytes:
    with open(ruta, "rb") as f:
        return pdf_to_image_bytes(f.read(), dpi=40)


# -------------------------------------------------------------------------
# DATOS
# -------------------------------------------------------------------------
def cargar_firmantes():
    return leer_json(FIRMANTES_FILE, [])


def guardar_firmantes(datos):
    escribir_json(FIRMANTES_FILE, datos)


def cargar_condecoraciones():
    data = leer_json(CONDECORACIONES_FILE, [])
    for item in data:
        item["config"] = {**DEFAULT_CONFIG, **item.get("config", {})}
        item.setdefault("id", slug(item.get("nombre", "")))
    return data


def guardar_condecoraciones(datos):
    escribir_json(CONDECORACIONES_FILE, datos)


def cargar_codigos():
    return leer_json(CODIGOS_FILE, [])


def guardar_codigos(datos):
    escribir_json(CODIGOS_FILE, datos)


def cargar_registro():
    return leer_json(REGISTRO_FILE, [])


def guardar_registro(datos):
    escribir_json(REGISTRO_FILE, datos)


def cargar_usuarios():
    return leer_json(USUARIOS_FILE, [])


def guardar_usuarios(datos):
    escribir_json(USUARIOS_FILE, datos)


# -------------------------------------------------------------------------
# CONTROL DE ACCESO
# -------------------------------------------------------------------------
def hash_clave(clave: str, sal: str = None):
    sal = sal or secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), bytes.fromhex(sal), 200_000).hex()
    return sal, h


def verificar_clave(clave: str, sal: str, hash_guardado: str) -> bool:
    try:
        return hmac.compare_digest(hash_clave(clave, sal)[1], hash_guardado)
    except Exception:
        return False


def admin_maestro():
    """Administrador principal definido en los Secrets de Streamlit ([admin] usuario / clave)."""
    try:
        sec = st.secrets["admin"]
        usuario = str(sec["usuario"]).strip().lower()
        clave = str(sec["clave"])
        if usuario and clave:
            return {"usuario": usuario, "clave": clave}
    except Exception:
        pass
    return None


def autenticar(usuario: str, clave: str):
    usuario = usuario.strip().lower()
    maestro = admin_maestro()
    if maestro and usuario == maestro["usuario"]:
        if hmac.compare_digest(clave.encode("utf-8"), maestro["clave"].encode("utf-8")):
            return {"usuario": usuario, "nombre": "Administrador principal", "rol": "admin", "maestro": True}
        return None
    for u in cargar_usuarios():
        if u.get("usuario") == usuario and u.get("activo", True) and u.get("rol") in ROLES:
            if verificar_clave(clave, u.get("sal", ""), u.get("hash", "")):
                return {"usuario": usuario, "nombre": u.get("nombre") or usuario, "rol": u["rol"], "maestro": False}
    return None


def sesion_vigente() -> bool:
    """Revalida la sesión en cada carga: si el usuario fue desactivado o cambió de nivel, se aplica al instante."""
    ses = st.session_state.get("sesion")
    if not ses:
        return False
    if ses.get("maestro"):
        m = admin_maestro()
        return bool(m) and m["usuario"] == ses["usuario"]
    u = next((x for x in cargar_usuarios() if x.get("usuario") == ses["usuario"]), None)
    if not u or not u.get("activo", True) or u.get("rol") not in ROLES:
        return False
    ses["rol"] = u["rol"]
    ses["nombre"] = u.get("nombre") or u["usuario"]
    return True


def puede(permiso: str) -> bool:
    ses = st.session_state.get("sesion")
    return bool(ses) and permiso in ROLES.get(ses["rol"], {}).get("permisos", set())


def usuario_actual() -> str:
    return st.session_state.get("sesion", {}).get("usuario", "")


def validar_nueva_clave(clave: str, repetir: str):
    if len(clave) < 8:
        return "La contraseña debe tener al menos 8 caracteres."
    if clave != repetir:
        return "Las contraseñas no coinciden."
    return None


def cerrar_sesion():
    for k in list(st.session_state.keys()):
        del st.session_state[k]


def pantalla_acceso():
    usuarios = cargar_usuarios()
    maestro = admin_maestro()
    _, centro, _ = st.columns([1, 1.3, 1])
    with centro:
        if not usuarios and not maestro:
            st.subheader("🔐 Configuración inicial")
            st.info("No hay usuarios todavía. Crea la cuenta del administrador principal.")
            with st.form("form_admin_inicial"):
                usuario = st.text_input("Usuario:", value="admin").strip().lower()
                nombre = st.text_input("Nombre completo:")
                clave = st.text_input("Contraseña (mínimo 8 caracteres):", type="password")
                repetir = st.text_input("Repetir contraseña:", type="password")
                if st.form_submit_button("Crear administrador", type="primary"):
                    error = validar_nueva_clave(clave, repetir)
                    if not re.fullmatch(r"[a-z0-9._-]{3,30}", usuario):
                        st.error("El usuario debe tener de 3 a 30 caracteres: letras, números, punto, guion o guion bajo.")
                    elif error:
                        st.error(error)
                    else:
                        sal, h = hash_clave(clave)
                        guardar_usuarios([{
                            "usuario": usuario, "nombre": nombre.strip() or usuario, "rol": "admin",
                            "activo": True, "sal": sal, "hash": h,
                            "creado": ahora().strftime("%Y-%m-%d %H:%M:%S"), "creado_por": "configuración inicial",
                        }])
                        st.session_state["sesion"] = {
                            "usuario": usuario, "nombre": nombre.strip() or usuario, "rol": "admin", "maestro": False,
                        }
                        st.rerun()
        else:
            st.subheader("🔐 Iniciar sesión")
            with st.form("form_login"):
                usuario = st.text_input("Usuario:")
                clave = st.text_input("Contraseña:", type="password")
                if st.form_submit_button("Entrar", type="primary"):
                    ses = autenticar(usuario, clave)
                    if ses:
                        st.session_state["sesion"] = ses
                        st.rerun()
                    else:
                        time.sleep(1.5)  # frena intentos repetidos
                        st.error("Usuario o contraseña incorrectos, o cuenta desactivada.")
    st.stop()


def barra_lateral():
    ses = st.session_state["sesion"]
    with st.sidebar:
        st.markdown(f"### 👤 {ses['nombre']}")
        st.caption(f"Usuario: `{ses['usuario']}` · Nivel: **{ROLES[ses['rol']]['nombre']}**")
        if st.button("Cerrar sesión"):
            cerrar_sesion()
            st.rerun()
        if not ses.get("maestro"):
            with st.expander("Cambiar mi contraseña"):
                with st.form("form_mi_clave", clear_on_submit=True):
                    actual = st.text_input("Contraseña actual:", type="password")
                    nueva = st.text_input("Nueva contraseña:", type="password")
                    repetir = st.text_input("Repetir nueva contraseña:", type="password")
                    if st.form_submit_button("Cambiar"):
                        usuarios = cargar_usuarios()
                        u = next((x for x in usuarios if x["usuario"] == ses["usuario"]), None)
                        error = validar_nueva_clave(nueva, repetir)
                        if not u or not verificar_clave(actual, u.get("sal", ""), u.get("hash", "")):
                            st.error("La contraseña actual no es correcta.")
                        elif error:
                            st.error(error)
                        else:
                            u["sal"], u["hash"] = hash_clave(nueva)
                            guardar_usuarios(usuarios)
                            st.success("Contraseña actualizada.")
        else:
            st.caption("La contraseña del administrador principal se cambia en los Secrets de Streamlit.")


# -------------------------------------------------------------------------
# CÓDIGOS, PREFIJOS Y REGISTRO
# -------------------------------------------------------------------------
PALABRAS_VACIAS = {"de", "en", "la", "el", "del", "y", "al", "a"}


def palabras_clave(texto) -> set:
    return {w for w in re.findall(r"[a-z0-9]+", normalizar(texto)) if w not in PALABRAS_VACIAS}


def sugerir_prefijo(nombre_condec: str, codigos: list):
    """Elige el prefijo cuya descripción se parece MÁS al nombre (no el primero que coincida)."""
    objetivo = palabras_clave(nombre_condec)
    mejor, mejor_puntaje = None, None
    for c in codigos:
        desc = palabras_clave(c.get("descripcion", ""))
        comunes = len(objetivo & desc)
        if comunes < 2:
            continue
        puntaje = (comunes, -len(objetivo ^ desc))
        if mejor_puntaje is None or puntaje > mejor_puntaje:
            mejor, mejor_puntaje = c["prefijo"], puntaje
    return mejor


def prefijo_de(condec: dict, codigos: list):
    """Prefijo asignado al formato; si no tiene, el más parecido por nombre."""
    pref = condec.get("prefijo")
    if pref and any(c["prefijo"] == pref for c in codigos):
        return pref
    return sugerir_prefijo(condec.get("nombre", ""), codigos)


def buscar_condecoracion(texto: str, condecoraciones: list):
    t = normalizar(texto)
    if not t:
        return None
    for c in condecoraciones:
        if normalizar(c["nombre"]) == t or c.get("id") == slug(texto):
            return c
    candidatos = [c for c in condecoraciones if t in normalizar(c["nombre"]) or normalizar(c["nombre"]) in t]
    return candidatos[0] if len(candidatos) == 1 else None


def siguiente_correlativo(registro: list, prefijo: str, ano: str, mes: str) -> int:
    base = f"{prefijo}-{ano}-{mes}-"
    mayor = 0
    for r in registro:
        cod = str(r.get("codigo", ""))
        if cod.startswith(base):
            m = re.match(r"(\d+)", cod[len(base):])
            if m:
                mayor = max(mayor, int(m.group(1)))
    return mayor + 1


def codigo_ocupado(registro: list, codigo: str, cedula_fmt: str) -> bool:
    """True si el código ya pertenece a OTRA persona."""
    return any(r.get("codigo") == codigo and r.get("cedula") != cedula_fmt for r in registro)


def registrar_emision(nombre, cedula, condec, codigo, lugar_fecha, firmante_izq, firmante_der, parrafo="") -> str:
    """Devuelve 'nuevo', 'actualizado' (mismo código y cédula: se corrige el asiento) u 'ocupado'."""
    reg = cargar_registro()
    cedula_fmt = formatear_cedula_vzla(cedula)
    codigo = codigo.strip()
    marca = ahora().strftime("%Y-%m-%d %H:%M:%S")
    datos = {
        "codigo": codigo,
        "condecoracion": condec["nombre"].strip(),
        "condecoracion_id": condec.get("id", ""),
        "nombre": nombre.strip().upper(),
        "cedula": cedula_fmt,
        "lugar_fecha": lugar_fecha.strip(),
        "parrafo": (parrafo or "").strip(),
        "firmante_izq": firmante_izq.get("nombre", "") if firmante_izq else "",
        "cargo_izq": firmante_izq.get("cargo", "") if firmante_izq else "",
        "firmante_der": firmante_der.get("nombre", "") if firmante_der else "",
        "cargo_der": firmante_der.get("cargo", "") if firmante_der else "",
    }
    for r in reg:
        if r.get("codigo") == codigo:
            if r.get("cedula") != cedula_fmt:
                return "ocupado"
            r.update(datos)
            r["fecha_actualizacion"] = marca
            r["actualizado_por"] = usuario_actual()
            guardar_registro(reg)
            return "actualizado"
    nuevo = {
        "id": f"{ahora().strftime('%Y%m%d%H%M%S')}_{secrets.token_hex(3)}",
        "fecha_registro": marca,
        "emitido_por": usuario_actual(),
        **datos,
    }
    reg.insert(0, nuevo)
    guardar_registro(reg)
    return "nuevo"


# -------------------------------------------------------------------------
# LECTURA DE LOTES (EXCEL / CSV)
# -------------------------------------------------------------------------
def normalizar_columna(col_nombre) -> str:
    texto = sin_acentos(str(col_nombre)).lower().replace(".", "")
    palabras = set(re.findall(r"[a-z0-9]+", texto))
    if palabras & {"condecoracion", "distincion", "premio", "reconocimiento"}:
        return "Condecoracion"
    if palabras & {"cedula", "ci", "documento", "identidad"}:
        return "Cedula"
    if palabras & {"correlativo", "numero", "nro", "num"}:
        return "Correlativo"
    if palabras & {"parrafo", "mensaje", "motivo", "dedicatoria", "texto"}:
        return "Parrafo"
    if palabras & {"nombre", "nombres", "galardonado", "beneficiario"}:
        return "Nombre"
    return str(col_nombre).strip().title()


def _fila_a_dict(encabezados, valores):
    fila = {}
    for i, h in enumerate(encabezados):
        if h is None:
            continue
        val = valores[i] if i < len(valores) else None
        if isinstance(val, float) and val.is_integer():
            val = int(val)
        txt = str(val).strip() if val is not None else ""
        if h in fila and fila[h]:
            continue  # ante encabezados repetidos, se conserva el primero con dato
        fila[h] = txt
    return fila


def leer_archivo_lote(uploaded_file):
    filas = []
    if uploaded_file.name.lower().endswith(".csv"):
        contenido = uploaded_file.getvalue().decode("utf-8-sig", errors="ignore")
        primera = contenido.splitlines()[0] if contenido.strip() else ""
        delimitador = ";" if primera.count(";") > primera.count(",") else ","
        filas_crudas = list(csv.reader(io.StringIO(contenido), delimiter=delimitador))
    else:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True, read_only=True)
        filas_crudas = [list(r) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
    if not filas_crudas:
        return filas
    encabezados = [
        normalizar_columna(h) if h is not None and str(h).strip() else None
        for h in filas_crudas[0]
    ]
    for r in filas_crudas[1:]:
        if any(cell is not None and str(cell).strip() for cell in r):
            filas.append(_fila_a_dict(encabezados, r))
    return filas


def crear_excel_ejemplo(condecoraciones):
    nombres = [c["nombre"] for c in condecoraciones] or ["Barra Espíritu Scout en Bronce"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Galardonados"
    ws.append(["Nombre", "Cedula", "Condecoracion", "Correlativo", "Parrafo"])
    ws.append(["CARLOS ALBERTO ESPINOZA", "33792612", nombres[0], "", ""])
    ws.append(["MARIA ELENA CHACON", "28.450.119", nombres[-1], "",
              "Por su invaluable aporte y dedicación constante al servicio de la comunidad scout."])
    ws.append(["JOSE LUIS URDANETA", "31902833", "", "", ""])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def exportar_registro_excel(registros):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Libro_Oficial_ASV"
    ws.append([
        "N° Registro", "Fecha y Hora", "Código de Control", "Condecoración", "Galardonado", "Cédula",
        "Lugar y Fecha", "Firmante Izquierdo", "Cargo Izq", "Firmante Derecho", "Cargo Der", "Emitido por",
    ])
    for idx, r in enumerate(registros, start=1):
        ws.append([
            idx, r.get("fecha_registro", ""), r.get("codigo", ""), r.get("condecoracion", ""),
            r.get("nombre", ""), r.get("cedula", ""), r.get("lugar_fecha", ""),
            r.get("firmante_izq", ""), r.get("cargo_izq", ""),
            r.get("firmante_der", ""), r.get("cargo_der", ""), r.get("emitido_por", ""),
        ])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# -------------------------------------------------------------------------
# ESTILOS Y ENCABEZADO
# -------------------------------------------------------------------------
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

# -------------------------------------------------------------------------
# PUERTA DE ACCESO
# -------------------------------------------------------------------------
if not sesion_vigente():
    st.session_state.pop("sesion", None)
    pantalla_acceso()  # termina con st.stop()

barra_lateral()
mostrar_avisos()


# -------------------------------------------------------------------------
# VISTA: EMITIR CERTIFICADO
# -------------------------------------------------------------------------
def selector_firmantes(firmantes, clave):
    nombres = ["(Ninguno)"] + [f["nombre"] for f in firmantes]
    c1, c2 = st.columns(2)
    with c1:
        izq = st.selectbox("Firmante izquierdo:", nombres, index=1 if len(nombres) > 1 else 0, key=f"{clave}_izq")
    with c2:
        der = st.selectbox("Firmante derecho:", nombres, index=2 if len(nombres) > 2 else 0, key=f"{clave}_der")
    return (
        next((f for f in firmantes if f["nombre"] == izq), None),
        next((f for f in firmantes if f["nombre"] == der), None),
    )


def emitir_individual(condecoraciones, firmantes, codigos):
    col_form, col_prev = st.columns([1, 1], gap="large")

    with col_form:
        st.subheader("1. Condecoración y galardonado")
        nombres_condec = [c["nombre"] for c in condecoraciones]
        sel = st.selectbox("Condecoración / formato:", nombres_condec, key="ind_condec")
        condec = next(c for c in condecoraciones if c["nombre"] == sel)

        nombre = st.text_input("Nombre y apellido:", key="ind_nombre",
                               help="Se imprime en mayúsculas, ajustado al ancho de la línea oficial.")
        cedula_txt = st.text_input("Cédula de identidad:", key="ind_cedula",
                                   help="Con o sin puntos. El sistema coloca los puntos automáticamente.")
        cedula = formatear_cedula_vzla(cedula_txt)
        if cedula:
            st.caption(f"Cédula formateada: **V- {cedula}**")

        parrafo_texto = ""
        if condec["config"].get("usa_parrafo"):
            parrafo_texto = st.text_area(
                "📝 Párrafo del certificado (va debajo de la cédula):",
                key=f"ind_parrafo_{condec['id']}", height=120,
                help="Se justifica y se ajusta el tamaño de letra automáticamente para que quepa "
                     "en el espacio calibrado para este formato.",
            )

        st.markdown("---")
        st.subheader("2. Fecha y código de registro")
        c_f1, c_f2 = st.columns(2)
        with c_f1:
            fecha = st.date_input("Fecha:", value=ahora().date(), key="ind_fecha")
        with c_f2:
            lugar_fecha = st.text_input("Lugar y fecha en el certificado:", value=texto_lugar_fecha(fecha))

        prefijos = [c["prefijo"] for c in codigos]
        descripciones = {c["prefijo"]: c.get("descripcion", "") for c in codigos}
        sugerido = prefijo_de(condec, codigos)
        idx_pref = prefijos.index(sugerido) if sugerido in prefijos else 0
        if not sugerido:
            st.warning("Este formato no tiene prefijo asignado ni uno parecido. Revisa el prefijo antes de emitir.")

        c_p, c_c = st.columns([2, 1])
        with c_p:
            prefijo = st.selectbox(
                "Prefijo oficial (Distrito-Grupo-Siglas):", prefijos, index=idx_pref,
                format_func=lambda p: f"{p} - {descripciones.get(p, '')}",
                key=f"ind_pref_{condec['id']}",
            )
        ano, mes = str(fecha.year), f"{fecha.month:02d}"
        siguiente = siguiente_correlativo(cargar_registro(), prefijo, ano, mes)
        with c_c:
            correlativo = st.number_input(
                "N° correlativo:", min_value=1, max_value=9999, value=min(siguiente, 9999), step=1,
                key=f"ind_corr_{prefijo}_{ano}_{mes}_{siguiente}",
                help="Se sugiere el siguiente número libre según el Libro de Registro.",
            )
        codigo_armado = f"{prefijo}-{ano}-{mes}-{int(correlativo):03d}"
        codigo = st.text_input("Código de control final (editable):", value=codigo_armado).strip()

        st.markdown("---")
        st.subheader("3. Firmantes oficiales")
        f_izq, f_der = selector_firmantes(firmantes, "ind")

        c_b1, c_b2 = st.columns(2)
        with c_b1:
            btn_prev = st.button("👁️ Vista previa (no registra)")
        with c_b2:
            btn_emitir = st.button("⚜️ Emitir y registrar", type="primary")

    with col_prev:
        st.subheader("Vista previa y descarga")

        if btn_prev or btn_emitir:
            ruta_template = os.path.join(PLANTILLAS_DIR, condec["archivo_pdf"])
            errores = []
            if not nombre.strip():
                errores.append("Escribe el nombre del galardonado.")
            if not os.path.exists(ruta_template):
                errores.append(f"El archivo de plantilla '{condec['archivo_pdf']}' no existe.")
            if btn_emitir and not codigo:
                errores.append("El código de control no puede estar vacío.")
            if btn_emitir and codigo and codigo_ocupado(cargar_registro(), codigo, cedula):
                errores.append(f"El código {codigo} ya está registrado a nombre de otra persona. Usa otro correlativo.")

            if errores:
                for e in errores:
                    st.error(e)
            else:
                with st.spinner("Procesando documento oficial..."):
                    pdf_bytes = generar_certificado_pdf(
                        template_path=ruta_template, nombre=nombre, cedula=cedula,
                        lugar_fecha=lugar_fecha, codigo_control=codigo,
                        firmante_izq=f_izq, firmante_der=f_der,
                        firmas_dir=FIRMAS_DIR, config=condec.get("config"),
                        parrafo=parrafo_texto,
                    )
                st.session_state["ind_pdf"] = pdf_bytes
                st.session_state["ind_archivo"] = f"Certificado_{slug(nombre)}_{codigo}.pdf"
                st.session_state["ind_emitido"] = False
                if btn_emitir:
                    estado = registrar_emision(nombre, cedula, condec, codigo, lugar_fecha, f_izq, f_der, parrafo_texto)
                    if estado == "ocupado":
                        st.error(f"El código {codigo} acaba de ser usado por otra emisión. Cambia el correlativo.")
                    else:
                        st.session_state["ind_emitido"] = True
                        if estado == "nuevo":
                            avisar(f"Certificado emitido y registrado con el código {codigo}.")
                        else:
                            avisar(f"El código {codigo} ya existía para esta cédula: se actualizó el asiento del Libro.", "info")
                        st.rerun()

        if "ind_pdf" in st.session_state:
            emitido = st.session_state.get("ind_emitido", False)
            st.image(
                pdf_to_image_bytes(st.session_state["ind_pdf"]),
                caption="Certificado emitido" if emitido else "BORRADOR: aún no registrado",
            )
            if emitido:
                st.download_button(
                    label="📥 Descargar certificado oficial en PDF",
                    data=st.session_state["ind_pdf"],
                    file_name=st.session_state.get("ind_archivo", "Certificado.pdf"),
                    mime="application/pdf",
                    type="primary",
                )
            else:
                st.info("Esto es un borrador. Pulsa «Emitir y registrar» para asentarlo en el Libro y habilitar la descarga.")
        else:
            st.info("Completa el formulario y pulsa «Vista previa» o «Emitir y registrar».")


def emitir_lote(condecoraciones, firmantes, codigos):
    st.subheader("⚡ Emisión masiva de certificados por lote")
    st.markdown("""
    El archivo puede llevar las columnas **Nombre**, **Cédula** (con o sin puntos) y, opcionalmente,
    **Condecoración**, **Correlativo** y **Parrafo** (para formatos como Agradecimiento).
    Si el correlativo va vacío, se asigna el siguiente número libre. Cada fila usa el prefijo de su
    propia condecoración.
    """)
    st.download_button(
        label="📄 Descargar plantilla Excel de ejemplo (.xlsx)",
        data=crear_excel_ejemplo(condecoraciones),
        file_name="plantilla_condecoraciones_scouts.xlsx",
        mime=MIME_XLSX,
    )
    st.markdown("---")

    col_cfg, col_file = st.columns([1, 1], gap="large")

    with col_cfg:
        st.markdown("#### Configuración del lote")
        nombres_condec = [c["nombre"] for c in condecoraciones]
        sel = st.selectbox("Condecoración para filas que no la indiquen:", nombres_condec, key="lote_condec")
        condec_defecto = next(c for c in condecoraciones if c["nombre"] == sel)

        c_f1, c_f2 = st.columns(2)
        with c_f1:
            fecha = st.date_input("Fecha de emisión:", value=ahora().date(), key="lote_fecha")
        with c_f2:
            lugar_fecha = st.text_input("Lugar y fecha:", value=texto_lugar_fecha(fecha))

        prefijos = [c["prefijo"] for c in codigos]
        descripciones = {c["prefijo"]: c.get("descripcion", "") for c in codigos}
        prefijo_respaldo = st.selectbox(
            "Prefijo para formatos sin prefijo asignado:", prefijos,
            format_func=lambda p: f"{p} - {descripciones.get(p, '')}", key="lote_pref",
            help="Solo se usa si la condecoración de una fila no tiene prefijo propio ni uno parecido.",
        )

        parrafo_defecto = ""
        if any(c.get("config", {}).get("usa_parrafo") for c in condecoraciones):
            st.markdown("#### Párrafo por defecto")
            parrafo_defecto = st.text_area(
                "Se usa en los formatos con párrafo si la fila no trae su propia columna «Parrafo»:",
                key="lote_parrafo_defecto", height=100,
            )

        st.markdown("#### Formato de impresión (hoja carta)")
        opcion_pag = st.radio(
            "Certificados por página:",
            ["1 por página (carta vertical)", "2 por página (carta horizontal)"],
            key="lote_por_pagina",
            help="Puedes cambiarlo después de generar el lote sin volver a emitir.",
        )
        por_pagina = 2 if opcion_pag.startswith("2") else 1
        linea_corte = False
        if por_pagina == 2:
            linea_corte = st.checkbox("Dibujar línea de corte al centro", value=True, key="lote_linea")
            st.caption("Cada certificado queda de unos 13 × 16,8 cm. Imprime al 100 %, sin «ajustar a la página».")

        st.markdown("#### Firmantes oficiales del lote")
        f_izq, f_der = selector_firmantes(firmantes, "lote")

    with col_file:
        st.markdown("#### Cargar archivo de galardonados")
        archivo = st.file_uploader("Archivo Excel (.xlsx) o CSV:", type=["xlsx", "csv"], key="lote_archivo")
        filas = []
        if archivo:
            try:
                filas = leer_archivo_lote(archivo)
            except Exception as e:
                st.error(f"No se pudo leer el archivo: {e}")
            if filas:
                for f in filas:
                    if "Cedula" in f:
                        f["Cedula"] = formatear_cedula_vzla(f["Cedula"])
                if "Nombre" not in filas[0]:
                    st.error("No se encontró una columna de nombres (Nombre, Galardonado o Beneficiario).")
                    filas = []
                else:
                    st.success(f"Archivo cargado: **{len(filas)} filas** detectadas.")
                    st.dataframe(filas[:10])
            elif archivo:
                st.error("No se encontraron filas con datos.")

        if filas and st.button("🚀 Procesar y generar lote completo", type="primary"):
            progreso = st.progress(0.0)
            estado_txt = st.empty()
            registro = cargar_registro()
            ano, mes = str(fecha.year), f"{fecha.month:02d}"
            contadores, usados = {}, set()
            pdfs, omitidos = [], []
            total = len(filas)

            for i, fila in enumerate(filas):
                n_fila = i + 2  # fila real en el Excel (la 1 es el encabezado)
                estado_txt.text(f"Procesando fila {i + 1} de {total}...")
                progreso.progress((i + 1) / total)
                try:
                    nom = str(fila.get("Nombre", "")).strip().upper()
                    ced = formatear_cedula_vzla(fila.get("Cedula", ""))
                    if not nom:
                        omitidos.append((n_fila, "(sin nombre)", "La fila no tiene nombre."))
                        continue

                    texto_condec = str(fila.get("Condecoracion", "")).strip()
                    if texto_condec:
                        condec = buscar_condecoracion(texto_condec, condecoraciones)
                        if not condec:
                            omitidos.append((n_fila, nom, f"Condecoración no reconocida: «{texto_condec}»."))
                            continue
                    else:
                        condec = condec_defecto

                    ruta_template = os.path.join(PLANTILLAS_DIR, condec["archivo_pdf"])
                    if not os.path.exists(ruta_template):
                        omitidos.append((n_fila, nom, f"Falta la plantilla '{condec['archivo_pdf']}'."))
                        continue

                    parrafo_fila = str(fila.get("Parrafo", "")).strip() or parrafo_defecto

                    prefijo = prefijo_de(condec, codigos) or prefijo_respaldo
                    base = f"{prefijo}-{ano}-{mes}-"
                    if base not in contadores:
                        contadores[base] = siguiente_correlativo(registro, prefijo, ano, mes)

                    num = None
                    corr_txt = str(fila.get("Correlativo", "")).strip()
                    if corr_txt:
                        try:
                            num = int(float(corr_txt))
                        except ValueError:
                            num = None
                    codigo = None
                    if num is None or num < 1:
                        # Si esta persona ya tiene esa condecoración este mes, se reutiliza su código
                        previo = next((r for r in registro if ced and r.get("cedula") == ced
                                       and str(r.get("codigo", "")).startswith(base)), None)
                        if previo:
                            codigo = previo["codigo"]
                        else:
                            num = contadores[base]
                    if codigo is None:
                        contadores[base] = max(contadores[base], num + 1)
                        codigo = f"{base}{num:03d}"

                    if codigo in usados or codigo_ocupado(registro, codigo, ced):
                        omitidos.append((n_fila, nom, f"El código {codigo} ya está en uso."))
                        continue

                    pdf_ind = generar_certificado_pdf(
                        template_path=ruta_template, nombre=nom, cedula=ced,
                        lugar_fecha=lugar_fecha, codigo_control=codigo,
                        firmante_izq=f_izq, firmante_der=f_der,
                        firmas_dir=FIRMAS_DIR, config=condec.get("config"),
                        parrafo=parrafo_fila,
                    )
                    if registrar_emision(nom, ced, condec, codigo, lugar_fecha, f_izq, f_der, parrafo_fila) == "ocupado":
                        omitidos.append((n_fila, nom, f"El código {codigo} ya está en uso."))
                        continue
                    usados.add(codigo)
                    pdfs.append((f"Certificado_{codigo}_{slug(nom)}.pdf", pdf_ind))
                except Exception as e:
                    omitidos.append((n_fila, str(fila.get("Nombre", "")), f"Error: {e}"))

            estado_txt.empty()
            progreso.empty()
            st.session_state["lote"] = {
                "pdfs": pdfs,
                "zip": crear_zip_certificados(pdfs) if pdfs else b"",
                "omitidos": omitidos,
                "fecha": fecha.strftime("%Y%m%d"),
                "clave_pdf": None,
                "pdf": b"",
            }

    # Resultados (persisten entre recargas; el formato de página se puede cambiar sin re-emitir)
    lote = st.session_state.get("lote")
    if lote:
        st.markdown("---")
        st.markdown("### 📥 Resultado del lote")
        n = len(lote["pdfs"])
        if n:
            clave = (por_pagina, linea_corte)
            if lote.get("clave_pdf") != clave:
                with st.spinner("Componiendo el PDF para imprimir..."):
                    lote["pdf"] = componer_paginas_carta(
                        [b for _, b in lote["pdfs"]], por_pagina=por_pagina, linea_corte=linea_corte,
                    )
                lote["clave_pdf"] = clave
            paginas = n if por_pagina == 1 else (n + 1) // 2
            st.success(f"{n} certificados emitidos y registrados.")
            c_d1, c_d2 = st.columns(2)
            with c_d1:
                st.download_button(
                    label=f"📄 PDF para imprimir: {por_pagina} por página ({paginas} hojas carta)",
                    data=lote["pdf"],
                    file_name=f"Lote_Condecoraciones_{lote['fecha']}_{por_pagina}x.pdf",
                    mime="application/pdf",
                    type="primary",
                    key="lote_dl_pdf",
                )
            with c_d2:
                st.download_button(
                    label="📦 ZIP con los PDF individuales",
                    data=lote["zip"],
                    file_name=f"Lote_Condecoraciones_{lote['fecha']}.zip",
                    mime="application/zip",
                    key="lote_dl_zip",
                )
        else:
            st.error("No se emitió ningún certificado.")
        if lote["omitidos"]:
            with st.expander(f"⚠️ {len(lote['omitidos'])} filas omitidas (no se emitieron)", expanded=True):
                for n_fila, nom, motivo in lote["omitidos"]:
                    st.markdown(f"- **Fila {n_fila}** · {nom}: {motivo}")
        if st.button("Cerrar resultado del lote"):
            st.session_state.pop("lote", None)
            st.rerun()


def vista_emitir():
    condecoraciones = cargar_condecoraciones()
    firmantes = cargar_firmantes()
    codigos = cargar_codigos()
    if not condecoraciones:
        st.warning("No hay formatos de condecoración cargados. Debe agregarlos un Gestor o Administrador.")
        return
    if not firmantes:
        st.warning("No hay firmantes registrados. Debe agregarlos un Gestor o Administrador.")
        return
    if not codigos:
        st.warning("No hay prefijos de código registrados. Debe agregarlos un Gestor o Administrador.")
        return
    modo = st.radio(
        "Modalidad de emisión:",
        ["Individual (uno a uno)", "⚡ Emisión masiva por lote (desde Excel)"],
        horizontal=True,
    )
    st.markdown("---")
    if modo.startswith("Individual"):
        emitir_individual(condecoraciones, firmantes, codigos)
    else:
        emitir_lote(condecoraciones, firmantes, codigos)


# -------------------------------------------------------------------------
# VISTA: LIBRO DE REGISTRO
# -------------------------------------------------------------------------
def vista_registro():
    st.subheader("📖 Libro de Registro Oficial de Condecoraciones y Distinciones")
    st.markdown("""
    > *Cumplimiento del **Artículo 35 del Reglamento Oficial**: Se hace constar el registro oficial de cada condecoración otorgada, antecedentes, fecha y número de certificado.*
    """)
    registro = cargar_registro()
    condecoraciones = cargar_condecoraciones()
    firmantes = cargar_firmantes()

    if not registro:
        st.info("Aún no se han emitido certificados. Cuando se emita el primero aparecerá aquí automáticamente.")
        return

    m1, m2, m3 = st.columns(3)
    m1.metric("Total de condecoraciones emitidas", len(registro))
    m2.metric("Última emisión", str(registro[0].get("fecha_registro", "-"))[:10])
    m3.metric("Condecoraciones diferentes", len(set(r.get("condecoracion", "") for r in registro)))
    st.markdown("---")

    col_filtro, col_exp = st.columns([3, 1])
    with col_filtro:
        busqueda = st.text_input("🔍 Buscar por nombre, cédula, código o condecoración:",
                                 placeholder="ej. Silverio, 11.289.602, NPB...")
    with col_exp:
        st.download_button(
            label="📥 Exportar a Excel (.xlsx)",
            data=exportar_registro_excel(registro),
            file_name=f"Libro_Registro_Condecoraciones_ASV_{ahora().strftime('%Y%m%d')}.xlsx",
            mime=MIME_XLSX,
            type="primary",
        )

    if busqueda.strip():
        filtro = normalizar(busqueda)
        mostrar = [
            r for r in registro
            if filtro in normalizar(r.get("nombre", ""))
            or filtro in normalizar(r.get("cedula", ""))
            or filtro.replace(".", "") in str(r.get("cedula", "")).replace(".", "")
            or filtro in normalizar(r.get("codigo", ""))
            or filtro in normalizar(r.get("condecoracion", ""))
        ]
    else:
        mostrar = registro

    st.markdown(f"### 📋 Listado de emisiones ({len(mostrar)})")
    for idx, item in enumerate(mostrar):
        clave = f"{idx}_{item.get('id', item.get('codigo', ''))}"
        nombre_copia = f"Copia_{slug(item.get('nombre', ''))}_{item.get('codigo', '')}.pdf"
        c_det, c_a1, c_a2 = st.columns([4, 1.5, 1])
        with c_det:
            st.markdown(f"**{item.get('nombre', '')}** — `C.I. {item.get('cedula', '')}`")
            emitido_por = f" | 👤 {item['emitido_por']}" if item.get("emitido_por") else ""
            st.caption(
                f"🏅 **{item.get('condecoracion', '')}** | 🏷️ Código: `{item.get('codigo', '')}` | "
                f"📅 {item.get('lugar_fecha', '')}{emitido_por}"
            )
            if item.get("parrafo"):
                with st.expander("Ver párrafo"):
                    st.write(item["parrafo"])
        with c_a1:
            if st.button("🔄 Generar copia PDF", key=f"reprint_{clave}"):
                condec = next((c for c in condecoraciones if c.get("id") == item.get("condecoracion_id")), None)
                if not condec:
                    condec = next((c for c in condecoraciones
                                   if normalizar(c["nombre"]) == normalizar(item.get("condecoracion", ""))), None)
                if not condec:
                    st.error(f"El formato «{item.get('condecoracion', '')}» ya no existe. No se puede generar la copia.")
                else:
                    ruta_tpl = os.path.join(PLANTILLAS_DIR, condec["archivo_pdf"])
                    if not os.path.exists(ruta_tpl):
                        st.error("Plantilla PDF no encontrada.")
                    else:
                        f_izq = next((f for f in firmantes if f["nombre"] == item.get("firmante_izq")), None)
                        f_der = next((f for f in firmantes if f["nombre"] == item.get("firmante_der")), None)
                        faltan = [n for n, o in ((item.get("firmante_izq"), f_izq), (item.get("firmante_der"), f_der)) if n and not o]
                        if faltan:
                            st.warning(f"Firmante(s) ya no registrados: {', '.join(faltan)}. La copia sale sin esa firma.")
                        st.session_state["copia_pdf_bytes"] = generar_certificado_pdf(
                            template_path=ruta_tpl, nombre=item.get("nombre", ""), cedula=item.get("cedula", ""),
                            lugar_fecha=item.get("lugar_fecha", ""), codigo_control=item.get("codigo", ""),
                            firmante_izq=f_izq, firmante_der=f_der,
                            firmas_dir=FIRMAS_DIR, config=condec.get("config"),
                            parrafo=item.get("parrafo", ""),
                        )
                        st.session_state["copia_pdf_nombre"] = nombre_copia
        with c_a2:
            if puede("anular") and st.button("🗑️ Borrar", key=f"del_reg_{clave}"):
                guardar_registro([r for r in registro if r is not item])
                avisar(f"Registro {item.get('codigo', '')} eliminado.")
                st.rerun()

        if st.session_state.get("copia_pdf_nombre") == nombre_copia:
            st.download_button(
                label=f"📥 Descargar copia PDF ({item.get('nombre', '')})",
                data=st.session_state["copia_pdf_bytes"],
                file_name=nombre_copia,
                mime="application/pdf",
                key=f"dl_copy_{clave}",
            )
        st.divider()

    if puede("limpiar"):
        with st.expander("🗑️ Mantenimiento general (solo Administrador)"):
            st.warning("Esto borra TODO el Libro de Registro. Descarga antes un respaldo o el Excel.")
            confirmar = st.checkbox("Entiendo que esta acción no se puede deshacer.", key="confirmar_limpiar")
            if st.button("Limpiar historial completo", disabled=not confirmar):
                guardar_registro([])
                avisar("Historial reiniciado.")
                st.rerun()


# -------------------------------------------------------------------------
# VISTA: CÓDIGOS DE CONTROL
# -------------------------------------------------------------------------
def vista_codigos():
    st.subheader("Configuración de prefijos para códigos de control")
    st.markdown("""
    Los códigos oficiales se arman con la siguiente estructura:
    - **Parte 1 (Distrito):** ej. `COQ`
    - **Parte 2 (Grupo Scout):** ej. `LCP`
    - **Parte 3 (Siglas de la condecoración):** ej. `NPB` (Nudo de Perseverancia en Bronce)
    - **Partes 4 y 5 (Año y mes):** se toman de la fecha del certificado (ej. `2026-04`)
    - **Parte 6 (Correlativo):** siguiente número libre de ese mes, o el que escribas (ej. `001`)
    """)
    codigos = cargar_codigos()

    if codigos:
        st.markdown(f"### Prefijos registrados ({len(codigos)})")
        busq = st.text_input("Filtrar por nombre o sigla:", placeholder="ej. Nudo, Barra, Oro, Bronce...")
        lista = codigos
        if busq.strip():
            b = normalizar(busq)
            lista = [c for c in codigos if b in normalizar(c["prefijo"]) or b in normalizar(c.get("descripcion", ""))]
        for idx, c in enumerate(lista):
            c1, c2, c3 = st.columns([3, 2, 1])
            with c1:
                st.markdown(f"**Prefijo:** `{c['prefijo']}`")
                st.caption(f"Descripción: {c.get('descripcion', 'Sin descripción')}")
            with c2:
                st.markdown(f"Distrito: **{c.get('distrito', '')}** | Grupo: **{c.get('grupo', '')}** | Siglas: **{c.get('siglas', '')}**")
            with c3:
                if st.button("🗑️ Eliminar", key=f"del_cod_{idx}_{c['prefijo']}"):
                    guardar_codigos([x for x in codigos if x is not c])
                    avisar(f"Prefijo {c['prefijo']} eliminado.")
                    st.rerun()
            st.divider()
    else:
        st.write("No hay prefijos registrados.")

    st.subheader("➕ Registrar nuevo prefijo")
    with st.form("form_nuevo_codigo", clear_on_submit=True):
        c_d, c_g, c_s = st.columns(3)
        with c_d:
            distrito = st.text_input("1. Distrito (ej. COQ):", max_chars=10)
        with c_g:
            grupo = st.text_input("2. Grupo Scout (ej. LCP):", max_chars=10)
        with c_s:
            siglas = st.text_input("3. Siglas de la condecoración (ej. NPB):", max_chars=10)
        desc = st.text_input("Descripción / nombre de la condecoración:", placeholder="ej. Nudo de Perseverancia en Bronce")
        if st.form_submit_button("Guardar prefijo", type="primary"):
            partes = [re.sub(r"[^A-Z0-9]", "", sin_acentos(p).upper()) for p in (distrito, grupo, siglas)]
            if not all(partes):
                st.error("Completa los 3 campos del código (solo letras y números).")
            else:
                prefijo = "-".join(partes)
                if any(c["prefijo"] == prefijo for c in codigos):
                    st.error(f"El prefijo {prefijo} ya existe.")
                else:
                    codigos.append({
                        "id": prefijo.lower().replace("-", "_"),
                        "distrito": partes[0], "grupo": partes[1], "siglas": partes[2],
                        "prefijo": prefijo, "descripcion": desc.strip() or prefijo,
                    })
                    guardar_codigos(codigos)
                    avisar(f"Prefijo {prefijo} registrado.")
                    st.rerun()


# -------------------------------------------------------------------------
# VISTA: FIRMANTES
# -------------------------------------------------------------------------
def vista_firmantes():
    st.subheader("Autoridades y firmantes registrados")
    st.info("Dirigentes y autoridades que firman las condecoraciones, con sus cargos y firmas digitalizadas.")
    firmantes = cargar_firmantes()

    if firmantes:
        for idx, f in enumerate(firmantes):
            c_info, c_img, c_del = st.columns([3, 2, 1])
            with c_info:
                st.markdown(f"**Nombre:** {f['nombre']}")
                st.markdown(f"**Cargo:** {f['cargo']}")
                st.markdown(f"**Institución:** {f.get('entidad', 'Asociación de Scouts de Venezuela')}")
            with c_img:
                ruta_img = os.path.join(FIRMAS_DIR, f.get("archivo_firma") or "_")
                if os.path.isfile(ruta_img):
                    st.image(ruta_img, width=150, caption="Firma digitalizada")
                else:
                    st.caption("Sin imagen de firma")
            with c_del:
                if st.button("🗑️ Eliminar", key=f"del_firm_{idx}_{f.get('id', '')}"):
                    guardar_firmantes([x for x in firmantes if x is not f])
                    avisar(f"Firmante «{f['nombre']}» eliminado.")
                    st.rerun()
            st.divider()
    else:
        st.write("No hay firmantes registrados.")

    st.subheader("➕ Agregar nuevo firmante")
    with st.form("form_nuevo_firmante", clear_on_submit=True):
        nombre = st.text_input("Nombre completo:")
        cargo = st.text_input("Cargo oficial:", placeholder="ej. JEFE DE GRUPO, COMISIONADO DISTRITAL")
        entidad = st.text_input("Institución / entidad:", value="Asociación de Scouts de Venezuela")
        archivo = st.file_uploader("Imagen de la firma (PNG con fondo transparente recomendado):", type=["png", "jpg", "jpeg"])
        if st.form_submit_button("Guardar firmante", type="primary"):
            if not nombre.strip() or not cargo.strip():
                st.error("Ingresa al menos el nombre y el cargo.")
            elif any(normalizar(f["nombre"]) == normalizar(nombre) for f in firmantes):
                st.error("Ya existe un firmante con ese nombre.")
            else:
                nombre_archivo = ""
                if archivo:
                    ext = os.path.splitext(archivo.name)[1].lower() or ".png"
                    nombre_archivo = f"{slug(nombre)}{ext}"
                    with open(os.path.join(FIRMAS_DIR, nombre_archivo), "wb") as f_out:
                        f_out.write(archivo.getbuffer())
                firmantes.append({
                    "id": slug(nombre), "nombre": nombre.strip(), "cargo": cargo.strip(),
                    "entidad": entidad.strip(), "archivo_firma": nombre_archivo,
                })
                guardar_firmantes(firmantes)
                avisar(f"Firmante «{nombre.strip()}» agregado.")
                st.rerun()


# -------------------------------------------------------------------------
# LIMPIEZA DE FORMATOS LLENOS (apoyo para la vista de formatos)
# -------------------------------------------------------------------------
def descartar_original(archivo_pdf: str):
    """Al subir una plantilla nueva, la copia original anterior ya no aplica."""
    ruta = os.path.join(ORIGINALES_DIR, archivo_pdf)
    if os.path.exists(ruta):
        os.remove(ruta)


def limpiar_plantilla(archivo_pdf: str, config: dict, profunda: bool = False) -> dict:
    """Limpia la plantilla partiendo siempre de su copia original (se puede repetir sin acumular daños)."""
    ruta = os.path.join(PLANTILLAS_DIR, archivo_pdf)
    original = os.path.join(ORIGINALES_DIR, archivo_pdf)
    if not os.path.exists(original):
        shutil.copy2(ruta, original)
    with open(original, "rb") as f:
        limpio, informe = limpiar_pdf_lleno(f.read(), config, profunda)
    with open(ruta, "wb") as f:
        f.write(limpio)
    return informe


def resumen_limpieza(nombre: str, informe: dict):
    partes = []
    if informe["textos"]:
        partes.append("texto eliminado: " + "; ".join(f"{zona}: «{txt}»" for zona, txt in informe["textos"]))
    if informe["imagenes"]:
        partes.append(f"{informe['imagenes']} firma(s) en imagen eliminada(s)")
    if partes:
        return "success", f"«{nombre}» limpiado. " + " | ".join(partes) + ". Revisa la vista grande del formato."
    if not informe["tiene_texto"]:
        return "warning", (f"«{nombre}»: el PDF no tiene texto editable (parece escaneado o aplanado). "
                           "Si está lleno, no se puede limpiar automáticamente; prueba la limpieza profunda "
                           "o sube la plantilla en blanco.")
    return "info", f"«{nombre}»: no se encontró nada que limpiar en las zonas variables."


# -------------------------------------------------------------------------
# VISTA: FORMATOS Y RESPALDO
# -------------------------------------------------------------------------
def vista_formatos():
    st.subheader("Catálogo de condecoraciones y formatos (PDF y Word)")
    st.info("Sube plantillas en PDF o Word (.docx), en blanco o ya llenas: la app borra el nombre, la cédula, "
            "las firmas, la fecha y el código, y deja el formato listo para emitir.")
    condec = cargar_condecoraciones()
    codigos = cargar_codigos()
    prefijos = [c["prefijo"] for c in codigos]
    descripciones = {c["prefijo"]: c.get("descripcion", "") for c in codigos}
    AUTO = "(Automático por nombre)"

    if len(condec) > 1:
        with st.expander("⚙️ Calibración por lote (copiar coordenadas a varios formatos)"):
            c_m1, c_m2 = st.columns([1, 2])
            with c_m1:
                modelo = st.selectbox("Copiar calibración DESDE:", [c["nombre"] for c in condec], key="modelo_lote")
            with c_m2:
                destinos = st.multiselect("Aplicar HACIA:", [c["nombre"] for c in condec if c["nombre"] != modelo], key="destinos_lote")
            if st.button("🚀 Aplicar calibración a los formatos seleccionados", type="primary"):
                modelo_obj = next((c for c in condec if c["nombre"] == modelo), None)
                if modelo_obj and destinos:
                    for c in condec:
                        if c["nombre"] in destinos:
                            c["config"] = dict(modelo_obj["config"])
                    guardar_condecoraciones(condec)
                    avisar(f"Calibración de «{modelo}» aplicada a {len(destinos)} formatos.")
                    st.rerun()
                else:
                    st.warning("Selecciona al menos un formato de destino.")

    st.markdown("---")

    if condec:
        st.markdown(f"### Formatos oficiales registrados ({len(condec)})")
        for idx, c in enumerate(condec):
            cfg = c["config"]
            k = f"{idx}_{c['id']}"
            ruta_actual = os.path.join(PLANTILLAS_DIR, c["archivo_pdf"])
            ruta_original = os.path.join(ORIGINALES_DIR, c["archivo_pdf"])

            c_info, c_prev, c_del = st.columns([3, 2, 1])
            with c_info:
                st.markdown(f"### {c['nombre']}")
                pref_actual = prefijo_de(c, codigos)
                origen = "asignado" if c.get("prefijo") in prefijos else "automático"
                extra = " · con párrafo" if cfg.get("usa_parrafo") else ""
                st.caption(f"Archivo: `{c['archivo_pdf']}` · Prefijo: `{pref_actual or 'sin prefijo'}` ({origen}){extra}")
            with c_prev:
                if os.path.exists(ruta_actual):
                    st.image(miniatura_pdf(ruta_actual, os.path.getmtime(ruta_actual)), width=110, caption="Miniatura")
                else:
                    st.caption("⚠️ Falta el archivo PDF")
            with c_del:
                if st.button("🗑️ Eliminar formato", key=f"del_condec_{k}"):
                    guardar_condecoraciones([x for x in condec if x is not c])
                    avisar(f"Formato «{c['nombre']}» eliminado.")
                    st.rerun()

            with st.expander(f"⚙️ Prefijo, calibración y limpieza de «{c['nombre']}»"):
                opciones = [AUTO] + prefijos
                idx_pref = opciones.index(c["prefijo"]) if c.get("prefijo") in prefijos else 0
                pref_sel = st.selectbox(
                    "Prefijo del código de control para este formato:", opciones, index=idx_pref,
                    format_func=lambda p: p if p == AUTO else f"{p} - {descripciones.get(p, '')}",
                    key=f"pref_{k}",
                )
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("**Nombre y cédula**")
                    n_nom_y = st.number_input("Altura Y del nombre:", value=float(cfg["nombre_y"]), step=1.0, key=f"ny_{k}")
                    n_nom_t = st.number_input("Tamaño máx. del nombre:", value=int(cfg["nombre_tamano_max"]), step=1, key=f"nt_{k}")
                    n_ced_x = st.number_input("Posición X de la cédula:", value=float(cfg["cedula_x"]), step=1.0, key=f"cx_{k}")
                    n_ced_y = st.number_input("Altura Y de la cédula:", value=float(cfg["cedula_y"]), step=0.5, key=f"cy_{k}")
                    n_ced_t = st.number_input("Tamaño de la cédula:", value=float(cfg["cedula_tamano"]), step=0.5, key=f"ct_{k}")
                with col2:
                    st.markdown("**Firmantes**")
                    n_fi_x = st.number_input("Centro X firmante izq.:", value=float(cfg["firmante_izq_x"]), step=2.0, key=f"fix_{k}")
                    n_fd_x = st.number_input("Centro X firmante der.:", value=float(cfg["firmante_der_x"]), step=2.0, key=f"fdx_{k}")
                    n_f_y = st.number_input("Altura Y de las firmas:", value=float(cfg["firmante_izq_y"]), step=2.0, key=f"fy_{k}")
                    n_f_a = st.number_input("Ancho de la firma:", value=float(cfg["ancho_firma"]), step=5.0, key=f"fa_{k}")
                with col3:
                    st.markdown("**Fecha y código de control**")
                    n_lf_x = st.number_input("Centro X de la fecha:", value=float(cfg["lugar_fecha_x"]), step=1.0, key=f"lx_{k}")
                    n_lf_y = st.number_input("Altura Y de la fecha:", value=float(cfg["lugar_fecha_y"]), step=1.0, key=f"ly_{k}")
                    n_cod_x = st.number_input("Centro X del código:", value=float(cfg["codigo_control_x"]), step=1.0, key=f"ccx_{k}")
                    n_cod_y = st.number_input("Altura Y del código:", value=float(cfg["codigo_control_y"]), step=0.5, key=f"ccy_{k}")
                    n_cod_t = st.number_input("Tamaño del código:", value=float(cfg["codigo_control_tamano"]), step=0.5, key=f"cct_{k}")

                st.markdown("---")
                st.markdown("**📝 Párrafo de agradecimiento (opcional)**")
                usa_parrafo = st.checkbox(
                    "Este formato tiene un párrafo debajo de la cédula",
                    value=bool(cfg.get("usa_parrafo", False)), key=f"usap_{k}",
                )
                if usa_parrafo:
                    cpa, cpb, cpc = st.columns(3)
                    with cpa:
                        n_par_sup = st.number_input(
                            "Altura Y superior (justo bajo la cédula):",
                            value=float(cfg.get("parrafo_y_sup", 400.0)), step=1.0, key=f"psup_{k}")
                        n_par_inf = st.number_input(
                            "Altura Y inferior (justo sobre el lema):",
                            value=float(cfg.get("parrafo_y_inf", 180.0)), step=1.0, key=f"pinf_{k}")
                    with cpb:
                        n_par_tmax = st.number_input(
                            "Tamaño máximo de letra:",
                            value=float(cfg.get("parrafo_tamano_max", 13.0)), step=0.5, key=f"ptmax_{k}")
                        n_par_tmin = st.number_input(
                            "Tamaño mínimo de letra:",
                            value=float(cfg.get("parrafo_tamano_min", 9.0)), step=0.5, key=f"ptmin_{k}")
                    with cpc:
                        n_par_ancho = st.number_input(
                            "Ancho (% del ancho de hoja, igual al del nombre):",
                            value=float(cfg.get("parrafo_ancho_pct", 0.84)) * 100, step=1.0, key=f"pancho_{k}",
                        ) / 100.0
                    st.caption("Escribe un texto de prueba largo en «Emitir Certificado» y revisa la vista previa "
                              "para terminar de ajustar estas alturas.")
                else:
                    n_par_sup = cfg.get("parrafo_y_sup", 400.0)
                    n_par_inf = cfg.get("parrafo_y_inf", 180.0)
                    n_par_tmax = cfg.get("parrafo_tamano_max", 13.0)
                    n_par_tmin = cfg.get("parrafo_tamano_min", 9.0)
                    n_par_ancho = cfg.get("parrafo_ancho_pct", 0.84)

                if st.button("💾 Guardar prefijo y calibración", key=f"save_cfg_{k}", type="primary"):
                    c["config"] = {
                        "nombre_y": n_nom_y, "nombre_tamano_max": n_nom_t,
                        "cedula_x": n_ced_x, "cedula_y": n_ced_y, "cedula_tamano": n_ced_t,
                        "lugar_fecha_x": n_lf_x, "lugar_fecha_y": n_lf_y,
                        "codigo_control_x": n_cod_x, "codigo_control_y": n_cod_y, "codigo_control_tamano": n_cod_t,
                        "firmante_izq_x": n_fi_x, "firmante_izq_y": n_f_y,
                        "firmante_der_x": n_fd_x, "firmante_der_y": n_f_y,
                        "ancho_firma": n_f_a,
                        "usa_parrafo": usa_parrafo,
                        "parrafo_y_sup": n_par_sup, "parrafo_y_inf": n_par_inf,
                        "parrafo_tamano_max": n_par_tmax, "parrafo_tamano_min": n_par_tmin,
                        "parrafo_ancho_pct": n_par_ancho,
                    }
                    if pref_sel == AUTO:
                        c.pop("prefijo", None)
                    else:
                        c["prefijo"] = pref_sel
                    guardar_condecoraciones(condec)
                    avisar(f"Configuración de «{c['nombre']}» guardada.")
                    st.rerun()

                st.markdown("---")
                st.markdown("**🧹 Limpieza de formato lleno**")
                st.caption("Borra lo que haya en las zonas de nombre, cédula, firmas, fecha, código y párrafo "
                           "(si aplica), según la calibración GUARDADA de este formato.")
                profunda = st.checkbox("Limpieza profunda (nombres convertidos a curvas)", key=f"prof_{k}")
                c_l1, c_l2 = st.columns(2)
                with c_l1:
                    if st.button("🧹 Limpiar este formato", key=f"limpiar_{k}"):
                        if not os.path.exists(ruta_actual):
                            st.error("Falta el archivo PDF de este formato.")
                        else:
                            try:
                                informe = limpiar_plantilla(c["archivo_pdf"], c["config"], profunda)
                                tipo_aviso, mensaje = resumen_limpieza(c["nombre"], informe)
                                avisar(mensaje, tipo_aviso)
                                st.rerun()
                            except Exception as e:
                                st.error(f"No se pudo limpiar: {e}")
                with c_l2:
                    if os.path.exists(ruta_original) and st.button("↩️ Restaurar original", key=f"rest_{k}"):
                        shutil.copy2(ruta_original, ruta_actual)
                        os.remove(ruta_original)
                        avisar(f"Se restauró el archivo original de «{c['nombre']}».")
                        st.rerun()
                if os.path.exists(ruta_actual) and st.checkbox("Ver el formato actual en grande", key=f"ver_{k}"):
                    with open(ruta_actual, "rb") as f_ver:
                        st.image(pdf_to_image_bytes(f_ver.read(), dpi=90), caption="Así quedó la plantilla")
            st.divider()
    else:
        st.write("No hay formatos registrados.")

    st.markdown("---")
    st.subheader("➕ Cargar formatos oficiales (PDF o Word .docx)")
    tipo = st.radio("Método de subida:", ["Subir un solo formato", "📦 Subir varios formatos a la vez"], horizontal=True)

    if tipo == "Subir un solo formato":
        with st.form("form_nueva_condec", clear_on_submit=True):
            nombre = st.text_input("Nombre de la distinción / condecoración:", placeholder="ej. Barra Espíritu Scout en Oro")
            archivo = st.file_uploader("Archivo oficial (.pdf o .docx):", type=["pdf", "docx"])
            textos = st.text_input(
                "Si el Word ya está rellenado, textos exactos a borrar (opcional, separados por coma):",
                placeholder="ej. JUAN PEREZ, 12.345.678, Maracaibo 10 de abril",
            )
            limpiar_lleno = st.checkbox(
                "Limpiar datos de un formato ya lleno (nombre, cédula, firmas, fecha y código)", value=True)
            limpieza_profunda = st.checkbox(
                "Limpieza profunda (si el nombre fue convertido a curvas)", value=False)
            if st.form_submit_button("Guardar y procesar formato", type="primary"):
                ident = slug(nombre)
                if not nombre.strip() or not archivo:
                    st.error("Ingresa el nombre y sube el archivo PDF o Word.")
                elif any(c["id"] == ident or normalizar(c["nombre"]) == normalizar(nombre) for c in condec):
                    st.error("Ya existe un formato con ese nombre. Elimínalo primero o usa otro nombre.")
                else:
                    try:
                        archivo_pdf = f"{ident}.pdf"
                        ruta_destino = os.path.join(PLANTILLAS_DIR, archivo_pdf)
                        if archivo.name.lower().endswith(".docx"):
                            borrar = [t.strip() for t in textos.split(",") if t.strip()]
                            with st.spinner("Limpiando el Word y convirtiéndolo a PDF..."):
                                limpiar_y_convertir_docx_a_pdf(archivo.getvalue(), ruta_destino, borrar)
                        else:
                            with open(ruta_destino, "wb") as f_out:
                                f_out.write(archivo.getbuffer())
                        descartar_original(archivo_pdf)
                        if limpiar_lleno:
                            informe = limpiar_plantilla(archivo_pdf, DEFAULT_CONFIG, limpieza_profunda)
                            tipo_aviso, mensaje = resumen_limpieza(nombre.strip(), informe)
                            avisar(mensaje, tipo_aviso)
                        condec.append({
                            "id": ident, "nombre": nombre.strip(),
                            "archivo_pdf": archivo_pdf, "config": DEFAULT_CONFIG.copy(),
                        })
                        guardar_condecoraciones(condec)
                        avisar(f"Formato «{nombre.strip()}» agregado. Si necesita párrafo, actívalo en su calibración.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al procesar el archivo: {e}")
    else:
        st.markdown("El nombre de cada formato se toma del nombre del archivo. Si el archivo ya existe, se reemplaza la plantilla y se conserva su calibración.")
        archivos = st.file_uploader("Archivos PDF o DOCX:", type=["pdf", "docx"], accept_multiple_files=True)
        textos_lote = st.text_input(
            "Textos de ejemplo a borrar en los Word (opcional, separados por coma):",
            placeholder="ej. NOMBRE DEL GALARDONADO, 00.000.000", key="limpiar_word_lote",
        )
        limpiar_lote = st.checkbox(
            "Limpiar datos de formatos ya llenos (nombre, cédula, firmas, fecha y código)",
            value=True, key="limpiar_lleno_lote")
        if archivos:
            st.info(f"**{len(archivos)} archivos** seleccionados.")
            if st.button("🚀 Subir, limpiar y registrar todo", type="primary"):
                agregados, actualizados = 0, 0
                borrar = [t.strip() for t in textos_lote.split(",") if t.strip()]
                with st.spinner("Procesando archivos..."):
                    for up in archivos:
                        try:
                            base = os.path.splitext(up.name)[0]
                            ident = slug(base)
                            archivo_pdf = f"{ident}.pdf"
                            ruta_destino = os.path.join(PLANTILLAS_DIR, archivo_pdf)
                            if up.name.lower().endswith(".docx"):
                                limpiar_y_convertir_docx_a_pdf(up.getvalue(), ruta_destino, borrar)
                            else:
                                with open(ruta_destino, "wb") as f_out:
                                    f_out.write(up.getbuffer())
                            existente = next((c for c in condec
                                              if c["archivo_pdf"] == archivo_pdf or c["id"] == ident), None)
                            descartar_original(archivo_pdf)
                            if limpiar_lote:
                                informe = limpiar_plantilla(
                                    archivo_pdf, existente["config"] if existente else DEFAULT_CONFIG)
                                tipo_aviso, mensaje = resumen_limpieza(up.name, informe)
                                avisar(mensaje, tipo_aviso)
                            if existente:
                                actualizados += 1
                            else:
                                condec.append({
                                    "id": ident,
                                    "nombre": base.replace("_", " ").replace("-", " ").strip().title(),
                                    "archivo_pdf": archivo_pdf, "config": DEFAULT_CONFIG.copy(),
                                })
                                agregados += 1
                        except Exception as e:
                            avisar(f"No se pudo procesar «{up.name}»: {e}", "warning")
                guardar_condecoraciones(condec)
                avisar(f"{agregados} formatos nuevos y {actualizados} plantillas reemplazadas.")
                st.rerun()

    if puede("respaldo"):
        st.markdown("---")
        st.subheader("🛡️ Respaldo y restauración total del sistema (ZIP)")
        st.markdown("""
        > **Protección para Streamlit Cloud:** descarga periódicamente una copia con las plantillas, firmas, códigos, usuarios y el Libro de Registro. Si el servidor se reinicia, restaura todo en segundos.
        """)
        c_r1, c_r2 = st.columns(2, gap="large")
        with c_r1:
            st.markdown("#### 1. Exportar copia de seguridad")
            if st.button("Preparar respaldo"):
                st.session_state["respaldo_zip"] = crear_respaldo_sistema_zip(PLANTILLAS_DIR, FIRMAS_DIR, DATOS_DIR)
                st.session_state["respaldo_nombre"] = f"Respaldo_Condecoraciones_ASV_{ahora().strftime('%Y%m%d_%H%M')}.zip"
            if "respaldo_zip" in st.session_state:
                st.download_button(
                    label="📥 Descargar respaldo completo (ZIP)",
                    data=st.session_state["respaldo_zip"],
                    file_name=st.session_state["respaldo_nombre"],
                    mime="application/zip",
                    type="primary",
                )
        with c_r2:
            st.markdown("#### 2. Restaurar desde respaldo")
            st.caption("Sobrescribe los datos actuales, incluidos los usuarios.")
            zip_subido = st.file_uploader("Archivo de respaldo (.zip):", type=["zip"], key="upload_zip_backup")
            if zip_subido and st.button("📤 Restaurar respaldo ahora", type="primary"):
                try:
                    total = restaurar_respaldo_sistema_zip(zip_subido.getvalue(), PLANTILLAS_DIR, FIRMAS_DIR, DATOS_DIR)
                    avisar(f"Sistema restaurado: {total} archivos recuperados.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al restaurar el respaldo: {e}")


# -------------------------------------------------------------------------
# VISTA: USUARIOS (solo Administrador)
# -------------------------------------------------------------------------
def hay_otro_admin(usuarios, excluir_usuario) -> bool:
    if admin_maestro():
        return True
    return any(u.get("rol") == "admin" and u.get("activo", True) and u["usuario"] != excluir_usuario for u in usuarios)


def vista_usuarios():
    st.subheader("👥 Usuarios y niveles de acceso")
    with st.expander("¿Qué puede hacer cada nivel?"):
        for rol in ROLES.values():
            st.markdown(f"- **{rol['nombre']}:** {rol['descripcion']}")

    usuarios = cargar_usuarios()
    yo = usuario_actual()
    maestro = admin_maestro()
    claves_rol = list(ROLES.keys())

    if maestro:
        st.info(f"Administrador principal definido en los Secrets: `{maestro['usuario']}`. No se puede editar ni borrar desde aquí.")
    else:
        st.warning("No hay administrador principal en los Secrets de Streamlit. Si el servidor se reinicia y se pierde "
                   "usuarios.json, cualquiera podría crear el primer administrador. Configura [admin] en Secrets.")

    if usuarios:
        st.markdown(f"### Usuarios registrados ({len(usuarios)})")
    for idx, u in enumerate(usuarios):
        es_yo = u["usuario"] == yo
        estado = "🟢 Activo" if u.get("activo", True) else "🔴 Desactivado"
        titulo = f"{u.get('nombre') or u['usuario']} · `{u['usuario']}` · {ROLES.get(u.get('rol'), {}).get('nombre', '?')} · {estado}"
        with st.expander(titulo + (" (tú)" if es_yo else "")):
            with st.form(f"form_usuario_{idx}_{u['usuario']}"):
                c1, c2 = st.columns(2)
                with c1:
                    nombre = st.text_input("Nombre completo:", value=u.get("nombre", ""))
                    rol = st.selectbox(
                        "Nivel de acceso:", claves_rol,
                        index=claves_rol.index(u["rol"]) if u.get("rol") in claves_rol else len(claves_rol) - 1,
                        format_func=lambda r: ROLES[r]["nombre"], disabled=es_yo,
                    )
                with c2:
                    activo = st.checkbox("Cuenta activa", value=u.get("activo", True), disabled=es_yo)
                    nueva = st.text_input("Nueva contraseña (vacío = no cambiar):", type="password")
                if st.form_submit_button("Guardar cambios", type="primary"):
                    pierde_admin = u.get("rol") == "admin" and u.get("activo", True) and (rol != "admin" or not activo)
                    if pierde_admin and not hay_otro_admin(usuarios, u["usuario"]):
                        st.error("Debe quedar al menos un administrador activo.")
                    elif nueva and len(nueva) < 8:
                        st.error("La contraseña debe tener al menos 8 caracteres.")
                    else:
                        u["nombre"] = nombre.strip() or u["usuario"]
                        if not es_yo:
                            u["rol"] = rol
                            u["activo"] = activo
                        if nueva:
                            u["sal"], u["hash"] = hash_clave(nueva)
                        guardar_usuarios(usuarios)
                        avisar(f"Usuario «{u['usuario']}» actualizado.")
                        st.rerun()
            st.caption(f"Creado: {u.get('creado', '-')} por {u.get('creado_por', '-')}")
            if not es_yo:
                confirmar = st.checkbox("Confirmo que quiero eliminar este usuario", key=f"conf_del_usr_{idx}")
                if st.button("🗑️ Eliminar usuario", key=f"del_usr_{idx}", disabled=not confirmar):
                    if u.get("rol") == "admin" and not hay_otro_admin(usuarios, u["usuario"]):
                        st.error("Debe quedar al menos un administrador activo.")
                    else:
                        guardar_usuarios([x for x in usuarios if x is not u])
                        avisar(f"Usuario «{u['usuario']}» eliminado.")
                        st.rerun()

    st.subheader("➕ Crear nuevo usuario")
    with st.form("form_nuevo_usuario", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            usuario = st.text_input("Usuario (para iniciar sesión):", placeholder="ej. carmen.chacon")
            nombre = st.text_input("Nombre completo:")
            rol = st.selectbox("Nivel de acceso:", claves_rol, index=claves_rol.index("emisor"),
                               format_func=lambda r: f"{ROLES[r]['nombre']}: {ROLES[r]['descripcion']}")
        with c2:
            clave = st.text_input("Contraseña inicial (mínimo 8 caracteres):", type="password")
            repetir = st.text_input("Repetir contraseña:", type="password")
        if st.form_submit_button("Crear usuario", type="primary"):
            usuario = usuario.strip().lower()
            error = validar_nueva_clave(clave, repetir)
            if not re.fullmatch(r"[a-z0-9._-]{3,30}", usuario):
                st.error("El usuario debe tener de 3 a 30 caracteres: letras sin acentos, números, punto, guion o guion bajo.")
            elif any(u["usuario"] == usuario for u in usuarios) or (maestro and maestro["usuario"] == usuario):
                st.error("Ese nombre de usuario ya existe.")
            elif error:
                st.error(error)
            else:
                sal, h = hash_clave(clave)
                usuarios.append({
                    "usuario": usuario, "nombre": nombre.strip() or usuario, "rol": rol, "activo": True,
                    "sal": sal, "hash": h,
                    "creado": ahora().strftime("%Y-%m-%d %H:%M:%S"), "creado_por": yo,
                })
                guardar_usuarios(usuarios)
                avisar(f"Usuario «{usuario}» creado con nivel {ROLES[rol]['nombre']}.")
                st.rerun()


# -------------------------------------------------------------------------
# PESTAÑAS SEGÚN EL NIVEL DEL USUARIO
# -------------------------------------------------------------------------
pestanas = []
if puede("emitir"):
    pestanas.append(("📜 Emitir Certificado", vista_emitir))
if puede("registro"):
    pestanas.append(("📖 Libro de Registro (Art. 35)", vista_registro))
if puede("codigos"):
    pestanas.append(("🏷️ Códigos de Control", vista_codigos))
if puede("firmantes"):
    pestanas.append(("✍️ Gestión de Firmantes", vista_firmantes))
if puede("formatos"):
    pestanas.append(("📁 Formatos y Respaldo", vista_formatos))
if puede("usuarios"):
    pestanas.append(("👥 Usuarios", vista_usuarios))

for contenedor, (_, vista) in zip(st.tabs([p[0] for p in pestanas]), pestanas):
    with contenedor:
        vista()

st.markdown("""
<div class="scout-footer">
    ⚜️ <strong>Asociación de Scouts de Venezuela</strong> • Fundada en 1913 • Reconocida por WOSM<br>
    <em>"Construir un mundo mejor"</em> • Sistema Oficial de Condecoraciones
</div>
""", unsafe_allow_html=True)
