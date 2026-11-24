import io
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import pypdf
import pymupdf

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
    
    # 1. Nombre del Galardonado (Centrado sobre la línea)
    c.setFont("Helvetica-Bold", 18)
    c.setFillColorRGB(0.1, 0.1, 0.1)
    c.drawCentredString(306, 467, nombre.upper())
    
    # 2. Cédula de Identidad (A la derecha de "C.I. V-")
    cedula_limpia = cedula.replace("V-", "").replace("v-", "").replace("C.I.", "").strip()
    c.setFont("Helvetica-Bold", 11)
    c.drawString(258, 441, cedula_limpia)
    
    # 3. Firmante Izquierdo
    if firmante_izq:
        x_izq = 160
        # Imagen de la firma
        img_firma = firmante_izq.get("archivo_firma")
        if img_firma:
            ruta_firma = os.path.join(firmas_dir, img_firma)
            if os.path.exists(ruta_firma):
                c.drawImage(ruta_firma, x_izq - 60, 140, width=120, height=55, mask='auto', preserveAspectRatio=True)
        # Textos del firmante
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x_izq, 130, firmante_izq.get("nombre", ""))
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x_izq, 118, firmante_izq.get("cargo", ""))
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(x_izq, 107, firmante_izq.get("entidad", ""))

    # 4. Firmante Derecho
    if firmante_der:
        x_der = 452
        # Imagen de la firma
        img_firma = firmante_der.get("archivo_firma")
        if img_firma:
            ruta_firma = os.path.join(firmas_dir, img_firma)
            if os.path.exists(ruta_firma):
                c.drawImage(ruta_firma, x_der - 60, 140, width=120, height=55, mask='auto', preserveAspectRatio=True)
        # Textos del firmante
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
        
    # 6. Código de Control / Registro ("Registrado bajo el número ... del control respectivo.")
    if codigo_control:
        c.setFont("Helvetica-Bold", 8)
        # El espacio en blanco va de 103pt a 234pt, su centro exacto es 168.5pt
        c.drawCentredString(168.5, 71, codigo_control)
        
    c.save()
    packet.seek(0)
    
    # Mezclar con el PDF original usando clone_from para compatibilidad con pypdf 7+
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
        nombre="Carlos Espinoza Nava",
        cedula="33.792.612",
        lugar_fecha="Maracaibo, 18 de abril de 2026",
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
    test_img_path = os.path.join(base_dir, "generados", "test_preview.png")
    with open(test_img_path, "wb") as f:
        f.write(img_bytes)
        
    print(f"Test generado con éxito en:\n- PDF: {test_pdf_path}\n- Imagen: {test_img_path}")
