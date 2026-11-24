import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
import pypdf
import pymupdf
from PIL import Image
import numpy as np

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
        # Identificar pixeles que no sean transparentes ni blanco absoluto
        ink = (alpha > 30) & (np.mean(rgb, axis=2) < 240)
        ys, xs = np.where(ink)
        if len(xs) > 0 and len(ys) > 0:
            x0 = max(0, xs.min() - 2)
            x1 = min(img.width, xs.max() + 2)
            y0 = max(0, ys.min() - 2)
            y1 = min(img.height, ys.max() + 2)
            cropped = img.crop((x0, y0, x1, y1))
            
            # Guardar en memoria o archivo temporal
            temp_path = os.path.join(os.path.dirname(ruta_imagen), f"_temp_crop_{os.path.basename(ruta_imagen)}")
            cropped.save(temp_path)
            return temp_path
    except Exception as e:
        print(f"Error al recortar firma: {e}")
    return ruta_imagen

def calcular_tamano_nombre(texto: str, ancho_linea: float = 515.0, porcentaje_objetivo: float = 0.88, max_tamano: int = 36, min_tamano: int = 18) -> int:
    """
    Calcula el tamaño de la fuente para que el nombre ocupe
    aproximadamente el 90% (88%-92%) de la línea oficial de 515pt.
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
    firmas_dir: str = ""
) -> bytes:
    """
    Superpone los textos y firmas sobre el PDF oficial de condecoración.
    Tamaño de página: 612 x 792 pt (Letter portrait)
    """
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    
    # 1. Nombre del Galardonado
    # La línea oficial mide 515pt (de X=50 a X=565).
    # Ajustamos la fuente para que ocupe aprox el 90% de la línea (max 36pt).
    nombre_txt = nombre.upper().strip()
    tam_nombre = calcular_tamano_nombre(nombre_txt, ancho_linea=515, porcentaje_objetivo=0.88, max_tamano=36, min_tamano=18)
    c.setFont("Helvetica-Bold", tam_nombre)
    c.setFillColorRGB(0.1, 0.1, 0.1)
    c.drawCentredString(306, 467, nombre_txt)
    
    # 2. Cédula de Identidad
    # La 'V-' en la plantilla mide 11.5pt de alto. Con fuente 16.5pt, los números
    # tienen exactamente la misma altura de la 'V-'. Su línea base está en Y=435.1.
    cedula_limpia = cedula.replace("V-", "").replace("v-", "").replace("C.I.", "").strip()
    c.setFont("Helvetica-Bold", 16.5)
    c.drawString(257, 435.1, cedula_limpia)
    
    # Ancho uniforme para ambas firmas (135 pt de tinta efectiva)
    ANCHO_FIRMA = 135
    
    # 3. Firmante Izquierdo
    if firmante_izq:
        x_izq = 160
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
                    138,
                    width=ANCHO_FIRMA,
                    height=h_calc,
                    mask='auto',
                    preserveAspectRatio=True
                )
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x_izq, 130, firmante_izq.get("nombre", ""))
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x_izq, 118, firmante_izq.get("cargo", ""))
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(x_izq, 107, firmante_izq.get("entidad", ""))

    # 4. Firmante Derecho
    if firmante_der:
        x_der = 452
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
                    138,
                    width=ANCHO_FIRMA,
                    height=h_calc,
                    mask='auto',
                    preserveAspectRatio=True
                )
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x_der, 130, firmante_der.get("nombre", ""))
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x_der, 118, firmante_der.get("cargo", ""))
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(x_der, 107, firmante_der.get("entidad", ""))
        
    # 5. Lugar y Fecha
    if lugar_fecha:
        c.setFont("Helvetica", 9)
        c.drawCentredString(306, 90, lugar_fecha)
        
    # 6. Código de Control / Registro
    # "Registrado bajo el número" termina en X=126.5pt
    # "del control respectivo." empieza en X=230pt
    # El centro del espacio en blanco es X=178.2pt, y la línea base de la frase es Y=65.5pt
    if codigo_control:
        cod_txt = codigo_control.strip()
        # Ajustamos el tamaño para que quepa perfectamente dentro del espacio sin tocar las palabras
        tam_cod = 7.5
        while tam_cod > 5.5 and stringWidth(cod_txt, "Helvetica-Bold", tam_cod) > 96:
            tam_cod -= 0.5
        c.setFont("Helvetica-Bold", tam_cod)
        c.drawCentredString(178.2, 65.5, cod_txt)
        
    c.save()
    packet.seek(0)
    
    # Mezclar con el PDF original
    writer = pypdf.PdfWriter(clone_from=template_path)
    overlay_pdf = pypdf.PdfReader(packet)
    writer.pages[0].merge_page(overlay_pdf.pages[0])
    
    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output.getvalue()

def pdf_to_image_bytes(pdf_bytes: bytes) -> bytes:
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=150)
    return pix.tobytes("png")

if __name__ == "__main__":
    base_dir = r"C:\Users\silve\.gemini\antigravity\scratch\certificados-app"
    template = os.path.join(base_dir, "plantillas", "barra_espiritu_bronce.pdf")
    firmas_d = os.path.join(base_dir, "firmas")
    
    pdf_out = generar_certificado_pdf(
        template_path=template,
        nombre="SILVERIO A. URDANETA S.",
        cedula="11.289.602",
        lugar_fecha="Maracaibo, 24 de abril de 2026",
        codigo_control="COQ-LCP-NPB-2026-04-001",
        firmante_izq={
            "nombre": "Carmen Chacón",
            "cargo": "JEFE DE GRUPO",
            "entidad": "Asociación de Scouts de Venezuela",
            "archivo_firma": "carmen_chacon.png"
        },
        firmante_der={
            "nombre": "Silverio Urdaneta",
            "cargo": "SUB JEFE DE GRUPO",
            "entidad": "Asociación de Scouts de Venezuela",
            "archivo_firma": "silverio_urdaneta.png"
        },
        firmas_dir=firmas_d
    )
    
    test_pdf_path = os.path.join(base_dir, "generados", "test_certificado.pdf")
    with open(test_pdf_path, "wb") as f:
        f.write(pdf_out)
        
    img_bytes = pdf_to_image_bytes(pdf_out)
    test_img_path = os.path.join(base_dir, "generados", "test_preview_adjusted.png")
    with open(test_img_path, "wb") as f:
        f.write(img_bytes)
        
    print(f"Test generado con éxito en:\n- PDF: {test_pdf_path}\n- Imagen: {test_img_path}")
