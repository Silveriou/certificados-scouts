import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from functools import lru_cache
from xml.sax.saxutils import escape as xml_escape

import docx
import numpy as np
import pymupdf
import pypdf
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph
from PIL import Image
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

CARTA_ANCHO = 612.0   # 8.5 pulgadas
CARTA_ALTO = 792.0    # 11 pulgadas
ALTO_MAX_FIRMA = 60.0

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
    "ancho_firma": 135.0,
    # Párrafo opcional (ej. certificados de Agradecimiento)
    "usa_parrafo": False,
    "parrafo_y_sup": 400.0,
    "parrafo_y_inf": 180.0,
    "parrafo_tamano_max": 13.0,
    "parrafo_tamano_min": 9.0,
    "parrafo_ancho_pct": 0.84,
}


def formatear_cedula_vzla(cedula_raw) -> str:
    """'11289602' o 'V-11.289.602' -> '11.289.602'"""
    if not cedula_raw:
        return ""
    solo_digitos = re.sub(r"\D", "", str(cedula_raw))
    if not solo_digitos:
        return str(cedula_raw).replace("V-", "").replace("v-", "").replace("C.I.", "").strip()
    return f"{int(solo_digitos):,}".replace(",", ".")


# -------------------------------------------------------------------------
# FIRMAS (recorte en memoria, sin archivos temporales)
# -------------------------------------------------------------------------
@lru_cache(maxsize=32)
def _firma_recortada(ruta: str, mtime: float):
    """
    Devuelve la firma como imagen RGBA recortada a la tinta.
    El fondo blanco (p. ej. de un JPG) se vuelve transparente.
    'mtime' solo sirve para invalidar la caché si el archivo cambia.
    """
    img = Image.open(ruta).convert("RGBA")
    arr = np.array(img)
    brillo = arr[:, :, :3].mean(axis=2)
    tinta = (arr[:, :, 3] > 30) & (brillo < 240)
    if not tinta.any():
        return img
    arr[brillo >= 240, 3] = 0
    ys, xs = np.where(tinta)
    x0 = max(0, int(xs.min()) - 2)
    x1 = min(img.width, int(xs.max()) + 3)
    y0 = max(0, int(ys.min()) - 2)
    y1 = min(img.height, int(ys.max()) + 3)
    return Image.fromarray(arr).crop((x0, y0, x1, y1))


def _dibujar_firmante(c, firmante: dict, x: float, y: float, ancho: float, firmas_dir: str):
    archivo = firmante.get("archivo_firma")
    if archivo:
        ruta = os.path.join(firmas_dir, archivo)
        if os.path.isfile(ruta):
            try:
                img = _firma_recortada(ruta, os.path.getmtime(ruta))
                proporcion = img.height / img.width
                w, h = ancho, ancho * proporcion
                if h > ALTO_MAX_FIRMA:
                    h = ALTO_MAX_FIRMA
                    w = h / proporcion
                c.drawImage(ImageReader(img), x - w / 2, y, width=w, height=h, mask="auto")
            except Exception as e:
                print(f"No se pudo dibujar la firma '{archivo}': {e}")
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(x, y - 8, firmante.get("nombre", ""))
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(x, y - 20, firmante.get("cargo", ""))
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(x, y - 31, firmante.get("entidad", ""))


# -------------------------------------------------------------------------
# PÁRRAFO JUSTIFICADO (para formatos como Agradecimiento)
# -------------------------------------------------------------------------
def _construir_parrafo(texto_html: str, ancho: float, tam: float, interlinea: float = 1.25):
    estilo = ParagraphStyle(
        name="certificado_parrafo",
        fontName="Helvetica",
        fontSize=tam,
        leading=tam * interlinea,
        alignment=TA_JUSTIFY,
    )
    p = Paragraph(texto_html, estilo)
    _, alto = p.wrap(ancho, 10000)
    return p, alto


def dibujar_parrafo_justificado(c, texto: str, page_w: float, y_sup: float, y_inf: float,
                                ancho_pct: float = 0.84, tam_max: float = 13.0, tam_min: float = 9.0) -> bool:
    """
    Dibuja un párrafo justificado, centrado verticalmente, dentro de la franja [y_inf, y_sup]
    (coordenadas PDF medidas desde abajo). Reduce el tamaño de letra hasta que quepa.
    Devuelve True si cupo sin llegar al tamaño mínimo.
    """
    texto = (texto or "").strip()
    if not texto or y_sup <= y_inf:
        return True

    ancho = page_w * ancho_pct
    alto_disponible = y_sup - y_inf
    texto_html = xml_escape(texto).replace("\n", "<br/>")

    tam = tam_max
    p, alto = _construir_parrafo(texto_html, ancho, tam)
    while alto > alto_disponible and tam > tam_min:
        tam -= 0.5
        p, alto = _construir_parrafo(texto_html, ancho, tam)

    x = (page_w - ancho) / 2.0
    y = y_inf + max(0.0, (alto_disponible - alto) / 2.0)
    p.drawOn(c, x, y)
    return alto <= alto_disponible


# -------------------------------------------------------------------------
# GENERACIÓN DEL CERTIFICADO
# -------------------------------------------------------------------------
def calcular_tamano_nombre(texto: str, ancho_linea: float = 515.0, porcentaje_objetivo: float = 0.88,
                           max_tamano: int = 36, min_tamano: int = 18) -> int:
    """Reduce la fuente hasta que el nombre quepa en el ancho objetivo de la línea."""
    ancho_deseado = ancho_linea * porcentaje_objetivo
    tamano = max_tamano
    while tamano > min_tamano:
        if stringWidth(texto.upper(), "Helvetica-Bold", tamano) <= ancho_deseado:
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
    config: dict = None,
    parrafo: str = "",
) -> bytes:
    """Superpone textos y firmas sobre la plantilla PDF oficial."""
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

    # 1. Nombre
    nombre_txt = (nombre or "").upper().strip()
    tam_nombre = calcular_tamano_nombre(
        nombre_txt,
        ancho_linea=page_w * 0.84,
        porcentaje_objetivo=0.88,
        max_tamano=int(cfg["nombre_tamano_max"]),
        min_tamano=18,
    )
    c.setFont("Helvetica-Bold", tam_nombre)
    c.setFillColorRGB(0.1, 0.1, 0.1)
    c.drawCentredString(centro_x, cfg["nombre_y"], nombre_txt)

    # 2. Cédula
    c.setFont("Helvetica-Bold", cfg["cedula_tamano"])
    c.drawString(cfg["cedula_x"], cfg["cedula_y"], formatear_cedula_vzla(cedula))

    # 3. Párrafo opcional (debajo de la cédula)
    if cfg.get("usa_parrafo") and parrafo:
        c.setFillColorRGB(0.1, 0.1, 0.1)
        dibujar_parrafo_justificado(
            c, parrafo, page_w,
            y_sup=cfg.get("parrafo_y_sup", 400.0),
            y_inf=cfg.get("parrafo_y_inf", 180.0),
            ancho_pct=cfg.get("parrafo_ancho_pct", 0.84),
            tam_max=cfg.get("parrafo_tamano_max", 13.0),
            tam_min=cfg.get("parrafo_tamano_min", 9.0),
        )

    # 4 y 5. Firmantes
    if firmante_izq:
        _dibujar_firmante(c, firmante_izq, cfg["firmante_izq_x"], cfg["firmante_izq_y"], cfg["ancho_firma"], firmas_dir)
    if firmante_der:
        _dibujar_firmante(c, firmante_der, cfg["firmante_der_x"], cfg["firmante_der_y"], cfg["ancho_firma"], firmas_dir)

    # 6. Lugar y fecha
    if lugar_fecha:
        c.setFont("Helvetica", 9)
        c.drawCentredString(cfg.get("lugar_fecha_x", centro_x), cfg["lugar_fecha_y"], lugar_fecha)

    # 7. Código de control
    if codigo_control:
        cod_txt = codigo_control.strip()
        tam_cod = float(cfg["codigo_control_tamano"])
        while tam_cod > 5.5 and stringWidth(cod_txt, "Helvetica-Bold", tam_cod) > 96:
            tam_cod -= 0.5
        c.setFont("Helvetica-Bold", tam_cod)
        c.drawCentredString(cfg["codigo_control_x"], cfg["codigo_control_y"], cod_txt)

    c.save()
    packet.seek(0)

    writer = pypdf.PdfWriter(clone_from=template_path)
    overlay_pdf = pypdf.PdfReader(packet)
    writer.pages[0].merge_page(overlay_pdf.pages[0])

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


# -------------------------------------------------------------------------
# COMPOSICIÓN DE LOTES EN HOJAS CARTA (1 o 2 certificados por página)
# -------------------------------------------------------------------------
def componer_paginas_carta(lista_pdf_bytes: list, por_pagina: int = 1,
                           margen: float = 14.0, linea_corte: bool = True) -> bytes:
    """
    Une certificados en un solo PDF listo para imprimir en tamaño carta.
    - por_pagina=1: carta vertical, un certificado a página completa.
    - por_pagina=2: carta horizontal, dos certificados lado a lado
      (reducidos proporcionalmente) con línea de corte opcional al centro.
    """
    salida = pymupdf.open()
    fuentes = []
    try:
        if por_pagina == 2:
            ancho, alto = CARTA_ALTO, CARTA_ANCHO
            mitad = ancho / 2.0
            for i in range(0, len(lista_pdf_bytes), 2):
                pagina = salida.new_page(width=ancho, height=alto)
                for j, datos in enumerate(lista_pdf_bytes[i:i + 2]):
                    src = pymupdf.open(stream=datos, filetype="pdf")
                    fuentes.append(src)
                    rect = pymupdf.Rect(j * mitad + margen, margen, (j + 1) * mitad - margen, alto - margen)
                    pagina.show_pdf_page(rect, src, 0)
                if linea_corte:
                    pagina.draw_line(
                        pymupdf.Point(mitad, 0), pymupdf.Point(mitad, alto),
                        color=(0.6, 0.6, 0.6), width=0.5, dashes="[4 4] 0",
                    )
        else:
            for datos in lista_pdf_bytes:
                src = pymupdf.open(stream=datos, filetype="pdf")
                fuentes.append(src)
                pagina = salida.new_page(width=CARTA_ANCHO, height=CARTA_ALTO)
                pagina.show_pdf_page(pymupdf.Rect(0, 0, CARTA_ANCHO, CARTA_ALTO), src, 0)
        return salida.tobytes(garbage=4, deflate=True)
    finally:
        for s in fuentes:
            s.close()
        salida.close()


def crear_zip_certificados(lista_certificados: list) -> bytes:
    """lista_certificados: lista de tuplas (nombre_archivo.pdf, bytes)"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for fname, data in lista_certificados:
            z.writestr(fname, data)
    return buf.getvalue()


def pdf_to_image_bytes(pdf_bytes: bytes, dpi: int = 150) -> bytes:
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        return doc[0].get_pixmap(dpi=dpi).tobytes("png")
    finally:
        doc.close()


# -------------------------------------------------------------------------
# LIMPIEZA DE FORMATOS PDF YA LLENOS
# -------------------------------------------------------------------------
def zonas_variables(page_w: float, page_h: float, config: dict = None) -> dict:
    """Rectángulos (coordenadas PyMuPDF) donde la app escribe los datos variables."""
    cfg = {**DEFAULT_CONFIG, **(config or {})}

    def r(x0, y_inf, x1, y_sup):
        # La calibración mide desde abajo; PyMuPDF mide desde arriba
        return pymupdf.Rect(max(0.0, x0), page_h - y_sup, min(page_w, x1), page_h - y_inf)

    mx = page_w * 0.05
    y_firma_min = min(cfg["firmante_izq_y"], cfg["firmante_der_y"])
    y_firma_max = max(cfg["firmante_izq_y"], cfg["firmante_der_y"])
    zonas = {
        "nombre": r(mx, cfg["nombre_y"] - 3, page_w - mx,
                    cfg["nombre_y"] + cfg["nombre_tamano_max"] + 8),
        "cédula": r(cfg["cedula_x"] + 1, cfg["cedula_y"] - 4, cfg["cedula_x"] + 190,
                    cfg["cedula_y"] + cfg["cedula_tamano"] + 2),
        "firmas y fecha": r(mx,
