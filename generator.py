import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from functools import lru_cache

import docx
import numpy as np
import pymupdf
import pypdf
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from PIL import Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

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

    # 3 y 4. Firmantes
    if firmante_izq:
        _dibujar_firmante(c, firmante_izq, cfg["firmante_izq_x"], cfg["firmante_izq_y"], cfg["ancho_firma"], firmas_dir)
    if firmante_der:
        _dibujar_firmante(c, firmante_der, cfg["firmante_der_x"], cfg["firmante_der_y"], cfg["ancho_firma"], firmas_dir)

    # 5. Lugar y fecha
    if lugar_fecha:
        c.setFont("Helvetica", 9)
        c.drawCentredString(cfg.get("lugar_fecha_x", centro_x), cfg["lugar_fecha_y"], lugar_fecha)

    # 6. Código de control
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
    return {
        "nombre": r(mx, cfg["nombre_y"] - 3, page_w - mx,
                    cfg["nombre_y"] + cfg["nombre_tamano_max"] + 8),
        "cédula": r(cfg["cedula_x"] + 1, cfg["cedula_y"] - 4, cfg["cedula_x"] + 190,
                    cfg["cedula_y"] + cfg["cedula_tamano"] + 2),
        "firmas y fecha": r(mx, min(cfg["lugar_fecha_y"] - 4, y_firma_min - 36), page_w - mx,
                            max(y_firma_max + ALTO_MAX_FIRMA + 4, cfg["lugar_fecha_y"] + 12)),
        "código": r(cfg["codigo_control_x"] - 49, cfg["codigo_control_y"] - 3,
                    cfg["codigo_control_x"] + 49, cfg["codigo_control_y"] + 10),
    }


def limpiar_pdf_lleno(pdf_bytes: bytes, config: dict = None, profunda: bool = False):
    """
    Quita de un formato ya lleno el nombre, la cédula, las firmas, la fecha y el código,
    sin tocar el fondo ni el texto fijo. Devuelve (pdf_limpio, informe).
    profunda=True también elimina trazos vectoriales totalmente contenidos en las zonas
    (para nombres convertidos a curvas).
    """
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = doc[0]
        zonas = zonas_variables(page.rect.width, page.rect.height, config)
        informe = {
            "textos": [],
            "imagenes": 0,
            "tiene_texto": bool(page.get_text("text").strip()),
        }

        # 1) Firmas colocadas como imagen (se ignoran imágenes grandes como el fondo)
        zona_firmas = zonas["firmas y fecha"]
        area_pagina = page.rect.get_area()
        borradas = set()
        for info in page.get_image_info(xrefs=True):
            xref = info.get("xref", 0)
            bbox = pymupdf.Rect(info["bbox"])
            area = bbox.get_area()
            if xref <= 0 or xref in borradas or area <= 0 or area > 0.25 * area_pagina:
                continue
            if bbox.intersects(zona_firmas) and (bbox & zona_firmas).get_area() / area >= 0.6:
                page.delete_image(xref)
                borradas.add(xref)
        informe["imagenes"] = len(borradas)

        # 2) Texto dentro de las zonas (sin pintar nada encima: el fondo queda igual)
        for nombre_zona, rect in zonas.items():
            texto = " ".join(page.get_text("text", clip=rect).split())
            if texto:
                informe["textos"].append((nombre_zona, texto))
            page.add_redact_annot(rect, fill=False, cross_out=False)
        page.apply_redactions(images=0, graphics=1 if profunda else 0)

        return doc.tobytes(garbage=3, deflate=True), informe
    finally:
        doc.close()


# -------------------------------------------------------------------------
# LIMPIEZA Y CONVERSIÓN DE PLANTILLAS WORD (.DOCX)
# -------------------------------------------------------------------------
PATRONES_RELLENO = [
    r"\{\{.*?\}\}",    # {{nombre}}
    r"\[[^\]]*\]",     # [Nombre del Galardonado]
    r"_{4,}",          # ______
    r"\.{5,}",         # ........
]


def _limpiar_texto(txt: str, extras: list) -> str:
    for pat in PATRONES_RELLENO:
        txt = re.sub(pat, " ", txt)
    for extra in extras:
        if extra and extra.strip():
            txt = re.sub(re.escape(extra.strip()), " ", txt, flags=re.IGNORECASE)
    return txt


def _limpiar_parrafo_docx(paragraph, extras: list):
    runs = [r for r in paragraph.runs if r.text]
    if not runs:
        return
    # 1) Limpieza run por run (conserva todo el formato)
    for run in runs:
        nuevo = _limpiar_texto(run.text, extras)
        if nuevo != run.text:
            run.text = nuevo
    # 2) Marcadores partidos entre varios runs: se consolida el párrafo
    completo = "".join(r.text for r in runs)
    limpio = _limpiar_texto(completo, extras)
    if limpio != completo:
        runs[0].text = limpio
        for r in runs[1:]:
            r.text = ""


def _parrafos(elemento, parent):
    """Todos los párrafos bajo un elemento: cuerpo, tablas y cuadros de texto."""
    for p_el in elemento.iter(qn("w:p")):
        yield Paragraph(p_el, parent)


def limpiar_y_convertir_docx_a_pdf(docx_bytes: bytes, ruta_pdf_salida: str, textos_extra_borrar: list = None):
    extras = textos_extra_borrar or []
    doc = docx.Document(io.BytesIO(docx_bytes))

    for p in _parrafos(doc.element.body, doc):
        _limpiar_parrafo_docx(p, extras)

    for section in doc.sections:
        for parte in (section.header, section.footer):
            if not parte.is_linked_to_previous:
                for p in _parrafos(parte._element, parte):
                    _limpiar_parrafo_docx(p, extras)

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_docx = os.path.join(tmpdir, "plantilla_limpia.docx")
        temp_pdf = os.path.join(tmpdir, "plantilla_limpia.pdf")
        doc.save(temp_docx)

        convertido = False
        soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice_bin:
            subprocess.run(
                [soffice_bin, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, temp_docx],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180,
            )
            if os.path.exists(temp_pdf):
                shutil.copy(temp_pdf, ruta_pdf_salida)
                convertido = True

        if not convertido:
            try:
                try:
                    import pythoncom
                    pythoncom.CoInitialize()
                except Exception:
                    pass
                from docx2pdf import convert
                convert(temp_docx, ruta_pdf_salida)
                convertido = os.path.exists(ruta_pdf_salida)
            except Exception as e:
                raise RuntimeError(f"No se pudo convertir el Word a PDF: {e}")

        if not convertido:
            raise RuntimeError("No se encontró LibreOffice ni Microsoft Word para convertir el archivo .docx a PDF.")


# -------------------------------------------------------------------------
# RESPALDO Y RESTAURACIÓN EN ZIP
# -------------------------------------------------------------------------
def crear_respaldo_sistema_zip(plantillas_dir: str, firmas_dir: str, datos_dir: str) -> bytes:
    buf = io.BytesIO()
    carpetas = {"plantillas": plantillas_dir, "firmas": firmas_dir, "datos": datos_dir}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre_carpeta, ruta_local in carpetas.items():
            if os.path.exists(ruta_local):
                for root, _, files in os.walk(ruta_local):
                    for file in files:
                        if file.startswith("_temp_") or file.endswith(".tmp"):
                            continue
                        full_path = os.path.join(root, file)
                        rel_path = os.path.relpath(full_path, ruta_local).replace("\\", "/")
                        z.write(full_path, f"{nombre_carpeta}/{rel_path}")
    return buf.getvalue()


def restaurar_respaldo_sistema_zip(zip_bytes: bytes, plantillas_dir: str, firmas_dir: str, datos_dir: str) -> int:
    """Restaura plantillas, firmas y datos. Ignora rutas que intenten salir de esas carpetas."""
    destinos = {"plantillas": plantillas_dir, "firmas": firmas_dir, "datos": datos_dir}
    restaurados = 0
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as z:
        for member in z.namelist():
            partes = [p for p in member.replace("\\", "/").split("/") if p]
            if len(partes) < 2 or partes[0] not in destinos or member.endswith("/"):
                continue
            if any(p in ("..", ".") or ":" in p for p in partes):
                continue
            base = os.path.realpath(destinos[partes[0]])
            ruta_final = os.path.realpath(os.path.join(base, *partes[1:]))
            try:
                if os.path.commonpath([base, ruta_final]) != base:
                    continue
            except ValueError:
                continue
            os.makedirs(os.path.dirname(ruta_final), exist_ok=True)
            with z.open(member) as source, open(ruta_final, "wb") as target:
                shutil.copyfileobj(source, target)
            restaurados += 1
    return restaurados
