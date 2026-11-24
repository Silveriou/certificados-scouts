import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
import pypdf
import pymupdf
from PIL import Image
import numpy as np
import docx

DEFAULT_CONFIG = {
    "nombre_y": 467.0,
    "nombre_tamano_max": 36,
    "cedula_x": 257.0,
    "cedula_y": 435.1,
    "cedula_tamano": 16.5,
    "lugar_fecha_x": 306.0,
    "lugar_fecha_y": 90.0,
    "codigo_control_x": 178.2,
    "codigo_control_y": 65.5,
    "codigo_control_tamano": 7.5,
    "firmante_izq_x": 160.0,
    "firmante_izq_y": 138.0,
    "firmante_der_x": 452.0,
    "firmante_der_y": 138.0,
    "ancho_firma": 135.0
}

def formatear_cedula_vzla(cedula_raw: str) -> str:
    """
    Formatea cualquier número de cédula venezolana con sus puntos reglamentarios.
    Ej: '11289602' o 'V-11.289.602' -> '11.289.602'
    """
    if not cedula_raw:
        return ""
    solo_digitos = re.sub(r"\D", "", str(cedula_raw))
    if not solo_digitos:
        return str(cedula_raw).replace("V-", "").replace("v-", "").replace("C.I.", "").strip()
    return f"{int(solo_digitos):,}".replace(",", ".")

def recortar_tinta_firma(ruta_imagen: str) -> str:
    """
    Recorta automáticamente los márgenes vacíos o blancos de la firma
    para que la tinta ocupe todo el ancho deseado.
    """
    try:
        img = Image.open(ruta_imagen).convert("RGBA")
        arr = np.array(img)
        alpha = arr[:, :, 3]
        rgb = arr[:, :, :3]
        ink = (alpha > 30) & (np.mean(rgb, axis=2) < 240)
        ys, xs = np.where(ink)
        if len(xs) > 0 and len(ys) > 0:
            x0 = max(0, xs.min() - 2)
            x1 = min(img.width, xs.max() + 2)
            y0 = max(0, ys.min() - 2)
            y1 = min(img.height, ys.max() + 2)
            cropped = img.crop((x0, y0, x1, y1))
            
            temp_path = os.path.join(os.path.dirname(ruta_imagen), f"_temp_crop_{os.path.basename(ruta_imagen)}")
            cropped.save(temp_path)
            return temp_path
    except Exception as e:
        print(f"Error al recortar firma: {e}")
    return ruta_imagen

def calcular_tamano_nombre(texto: str, ancho_linea: float = 515.0, porcentaje_objetivo: float = 0.88, max_tamano: int = 36, min_tamano: int = 18) -> int:
    """
    Calcula el tamaño de la fuente para que el nombre ocupe
    aproximadamente el 90% (88%-92%) de la línea oficial.
    """
    ancho_deseado = ancho_linea * porcentaje_objetivo
    tamano = max_tamano
    while tamano > min_tamano:
        w = stringWidth(texto.upper(), "Helvetica-Bold", tamano)
        if w <= ancho_deseado:
            break
        tamano -= 1
    return tamano

def generar_certificado_pdf(
    template_path: str,
    nombre: str,
    cedula: str,
    lugar_fecha: str,
    codigo_control: str,
    firmante_izq: dict = None,
    firmante_der: dict = None,
    firmas_dir: str = "",
    config: dict = None
) -> bytes:
    """
    Superpone los textos y firmas sobre el PDF oficial de condecoración
    utilizando las coordenadas calibradas para esa plantilla.
    """
    cfg = DEFAULT_CONFIG.copy()
    if config:
        cfg.update(config)

    base_pdf = pypdf.PdfReader(template_path)
    base_page = base_pdf.pages[0]
    page_w = float(base_page.mediabox.width)
    page_h = float(base_page.mediabox.height)
    centro_x = page_w / 2.0

    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_w, page_h))
    
    # 1. Nombre del Galardonado
    nombre_txt = nombre.upper().strip()
    ancho_linea = page_w * 0.84
    tam_max = int(cfg.get("nombre_tamano_max", 36))
    tam_nombre = calcular_tamano_nombre(nombre_txt, ancho_linea=ancho_linea, porcentaje_objetivo=0.88, max_tamano=tam_max, min_tamano=18)
    c.setFont("Helvetica-Bold", tam_nombre)
    c.setFillColorRGB(0.1, 0.1, 0.1)
    c.drawCentredString(centro_x, cfg.get("nombre_y", 467.0), nombre_txt)
    
    # 2. Cédula de Identidad (Formateada automáticamente con puntos)
    cedula_limpia = formatear_cedula_vzla(cedula)
    c.setFont("Helvetica-Bold", cfg.get("cedula_tamano", 16.5))
    c.drawString(cfg.get("cedula_x", 257.0), cfg.get("cedula_y", 435.1), cedula_limpia)
    
    # Ancho uniforme de firmas
    ANCHO_FIRMA = cfg.get("ancho_firma", 135.0)
    
    # 3. Firmante Izquierdo
    if firmante_izq:
        x_izq = cfg.get("firmante_izq_x", 160.0)
        y_izq = cfg.get("firmante_izq_y", 138.0)
        img_firma = firmante_izq.get("archivo_firma")
        if img_firma:
            ruta_firma = os.path.join(firmas_dir, img_firma)
            if os.path.exists(ruta_firma):
                ruta_recortada = recortar_tinta_firma(ruta_firma)
                with Image.open(ruta_recortada) as im:
                    ratio = im.height / im.width
                    h_calc = min(60, ANCHO_FIRMA * ratio)
                c.drawImage(
                    ruta_recortada,
                    x_izq - (ANCHO_FIRMA / 2),
                    y_izq,
                    width=ANCHO_FIRMA,
                    height=h_calc,
                    mask='auto',
                    preserveAspectRatio=True
                )
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x_izq, y_izq - 8, firmante_izq.get("nombre", ""))
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x_izq, y_izq - 20, firmante_izq.get("cargo", ""))
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(x_izq, y_izq - 31, firmante_izq.get("entidad", ""))

    # 4. Firmante Derecho
    if firmante_der:
        x_der = cfg.get("firmante_der_x", 452.0)
        y_der = cfg.get("firmante_der_y", 138.0)
        img_firma = firmante_der.get("archivo_firma")
        if img_firma:
            ruta_firma = os.path.join(firmas_dir, img_firma)
            if os.path.exists(ruta_firma):
                ruta_recortada = recortar_tinta_firma(ruta_firma)
                with Image.open(ruta_recortada) as im:
                    ratio = im.height / im.width
                    h_calc = min(60, ANCHO_FIRMA * ratio)
                c.drawImage(
                    ruta_recortada,
                    x_der - (ANCHO_FIRMA / 2),
                    y_der,
                    width=ANCHO_FIRMA,
                    height=h_calc,
                    mask='auto',
                    preserveAspectRatio=True
                )
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x_der, y_der - 8, firmante_der.get("nombre", ""))
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x_der, y_der - 20, firmante_der.get("cargo", ""))
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(x_der, y_der - 31, firmante_der.get("entidad", ""))
        
    # 5. Lugar y Fecha
    if lugar_fecha:
        c.setFont("Helvetica", 9)
        c.drawCentredString(cfg.get("lugar_fecha_x", centro_x), cfg.get("lugar_fecha_y", 90.0), lugar_fecha)
        
    # 6. Código de Control / Registro
    if codigo_control:
        cod_txt = codigo_control.strip()
        tam_cod = cfg.get("codigo_control_tamano", 7.5)
        while tam_cod > 5.5 and stringWidth(cod_txt, "Helvetica-Bold", tam_cod) > 96:
            tam_cod -= 0.5
        c.setFont("Helvetica-Bold", tam_cod)
        c.drawCentredString(cfg.get("codigo_control_x", 178.2), cfg.get("codigo_control_y", 65.5), cod_txt)
        
    c.save()
    packet.seek(0)
    
    writer = pypdf.PdfWriter(clone_from=template_path)
    overlay_pdf = pypdf.PdfReader(packet)
    writer.pages[0].merge_page(overlay_pdf.pages[0])
    
    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output.getvalue()

def unir_certificados_pdf(lista_pdf_bytes: list) -> bytes:
    """
    Une múltiples certificados PDF en un único archivo PDF continuo
    listo para imprimir en lote.
    """
    writer = pypdf.PdfWriter()
    for b in lista_pdf_bytes:
        reader = pypdf.PdfReader(io.BytesIO(b))
        writer.add_page(reader.pages[0])
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.getvalue()

def crear_zip_certificados(lista_certificados: list) -> bytes:
    """
    Empaqueta múltiples certificados PDF individuales en un archivo ZIP.
    lista_certificados: lista de tuplas (nombre_archivo.pdf, bytes)
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for fname, data in lista_certificados:
            z.writestr(fname, data)
    buf.seek(0)
    return buf.getvalue()

def pdf_to_image_bytes(pdf_bytes: bytes) -> bytes:
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=150)
    return pix.tobytes("png")

# -------------------------------------------------------------------------
# PROCESAMIENTO Y LIMPIEZA DE PLANTILLAS WORD (.DOCX)
# -------------------------------------------------------------------------
def _limpiar_parrafo_docx(paragraph, textos_extra_borrar: list):
    """
    Limpia marcadores de relleno, corchetes, llaves o textos de ejemplo
    manteniendo los estilos y el diseño intactos.
    """
    patrones_regex = [
        r"\{\{.*?\}\}",   # {{nombre}}, {{cedula}}
        r"\[.*?\]",         # [Nombre del Galardonado]
        r"<.*?>",             # <Nombre>
        r"_{4,}",             # Líneas de subrayado ______
        r"\.{5,}",           # Líneas de puntos ........
    ]
    
    for run in paragraph.runs:
        txt = run.text
        if not txt:
            continue
        for pat in patrones_regex:
            txt = re.sub(pat, " ", txt)
        for extra in textos_extra_borrar:
            if extra and extra.strip():
                txt = re.sub(re.escape(extra.strip()), " ", txt, flags=re.IGNORECASE)
        run.text = txt

def limpiar_y_convertir_docx_a_pdf(docx_bytes: bytes, ruta_pdf_salida: str, textos_extra_borrar: list = None):
    """
    1. Abre el .docx en memoria y elimina textos rellenados o marcadores.
    2. Convierte el documento limpio a PDF usando docx2pdf (Windows/Mac) o LibreOffice (Linux/Cloud).
    """
    if textos_extra_borrar is None:
        textos_extra_borrar = []

    doc = docx.Document(io.BytesIO(docx_bytes))

    # Limpiar párrafos principales
    for p in doc.paragraphs:
        _limpiar_parrafo_docx(p, textos_extra_borrar)

    # Limpiar celdas de tablas (muy común en diseños de certificados Word)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _limpiar_parrafo_docx(p, textos_extra_borrar)

    # Limpiar encabezados y pies de página
    for section in doc.sections:
        for p in section.header.paragraphs:
            _limpiar_parrafo_docx(p, textos_extra_borrar)
        for p in section.footer.paragraphs:
            _limpiar_parrafo_docx(p, textos_extra_borrar)

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_docx = os.path.join(tmpdir, "plantilla_limpia.docx")
        temp_pdf = os.path.join(tmpdir, "plantilla_limpia.pdf")
        doc.save(temp_docx)

        convertido = False
        # Intento 1: LibreOffice (ideal para Streamlit Cloud / Linux o si está instalado en Windows)
        soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice_bin:
            subprocess.run(
                [soffice_bin, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, temp_docx],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            if os.path.exists(temp_pdf):
                shutil.copy(temp_pdf, ruta_pdf_salida)
                convertido = True

        # Intento 2: docx2pdf (ideal en PC local con Microsoft Word instalado)
        if not convertido:
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except Exception:
                pass
            from docx2pdf import convert
            convert(temp_docx, ruta_pdf_salida)
            if os.path.exists(ruta_pdf_salida):
                convertido = True

        if not convertido:
            raise RuntimeError("No se encontró LibreOffice ni Microsoft Word para convertir el archivo .docx a PDF.")

# -------------------------------------------------------------------------
# SISTEMA DE RESPALDO Y RESTAURACIÓN EN ZIP
# -------------------------------------------------------------------------
def crear_respaldo_sistema_zip(plantillas_dir: str, firmas_dir: str, datos_dir: str) -> bytes:
    """
    Empaqueta todas las carpetas críticas (plantillas, firmas, datos) en un ZIP.
    """
    buf = io.BytesIO()
    carpetas = {
        "plantillas": plantillas_dir,
        "firmas": firmas_dir,
        "datos": datos_dir
    }
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre_carpeta, ruta_local in carpetas.items():
            if os.path.exists(ruta_local):
                for root, _, files in os.walk(ruta_local):
                    for file in files:
                        if file.startswith("_temp_"):
                            continue
                        full_path = os.path.join(root, file)
                        rel_path = os.path.relpath(full_path, ruta_local)
                        z.write(full_path, os.path.join(nombre_carpeta, rel_path))
    buf.seek(0)
    return buf.getvalue()

def restaurar_respaldo_sistema_zip(zip_bytes: bytes, plantillas_dir: str, firmas_dir: str, datos_dir: str) -> int:
    """
    Restaura las carpetas plantillas, firmas y datos desde un archivo ZIP de respaldo.
    """
    destinos = {
        "plantillas": plantillas_dir,
        "firmas": firmas_dir,
        "datos": datos_dir
    }
    archivos_restaurados = 0
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as z:
        for member in z.namelist():
            partes = member.replace("\\", "/").split("/")
            if len(partes) >= 2 and partes[0] in destinos and partes[-1]:
                carpeta_destino = destinos[partes[0]]
                nombre_archivo = "/".join(partes[1:])
                ruta_final = os.path.join(carpeta_destino, nombre_archivo)
                os.makedirs(os.path.dirname(ruta_final), exist_ok=True)
                with z.open(member) as source, open(ruta_final, "wb") as target:
                    shutil.copyfileobj(source, target)
                archivos_restaurados += 1
    return archivos_restaurados
