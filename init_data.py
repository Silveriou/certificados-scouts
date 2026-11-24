import json
import os

datos_dir = r"C:\Users\silve\.gemini\antigravity\scratch\certificados-app\datos"
os.makedirs(datos_dir, exist_ok=True)

firmantes_iniciales = [
    {
        "id": "carmen_chacon",
        "nombre": "Carmen Chacón",
        "cargo": "JEFE DE GRUPO",
        "entidad": "Asociación de Scouts de Venezuela",
        "archivo_firma": "carmen_chacon.png"
    },
    {
        "id": "silverio_urdaneta",
        "nombre": "Silverio Urdaneta",
        "cargo": "SUB JEFE DE GRUPO",
        "entidad": "Asociación de Scouts de Venezuela",
        "archivo_firma": "silverio_urdaneta.png"
    }
]

condecoraciones_iniciales = [
    {
        "id": "barra_espiritu_bronce",
        "nombre": "Barra Espíritu Scout en Bronce",
        "archivo_pdf": "barra_espiritu_bronce.pdf"
    }
]

with open(os.path.join(datos_dir, "firmantes.json"), "w", encoding="utf-8") as f:
    json.dump(firmantes_iniciales, f, indent=2, ensure_ascii=False)

with open(os.path.join(datos_dir, "condecoraciones.json"), "w", encoding="utf-8") as f:
    json.dump(condecoraciones_iniciales, f, indent=2, ensure_ascii=False)

print("Datos iniciales creados exitosamente")
