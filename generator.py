import io
import os
import zipfile
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
import pypdf
import pymupdf
from PIL import Image
import numpy as np

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
    
    # 2. Cédula de Identidad
    cedula_limpia = cedula.replace("V-", "").replace("v-", "").replace("C.I.", "").strip()
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
