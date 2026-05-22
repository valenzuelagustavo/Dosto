import os
import pdfplumber

# Rutas personalizadas exactas
CARPETA_PDFS = r"G:\Programas\Dosto\Facturas ejemplos"
ARCHIVO_SALIDA = r"G:\Programas\Dosto\formatos_crudos.txt"

if not os.path.exists(CARPETA_PDFS):
    print(f"-> Error: No se encontró la carpeta de origen: {CARPETA_PDFS}")
    print("Por favor, verificá que el nombre de la carpeta esté bien escrito.")
    exit()

archivos = [f for f in os.listdir(CARPETA_PDFS) if f.lower().endswith('.pdf')]

if not archivos:
    print(f"-> No encontré ningún archivo PDF dentro de: {CARPETA_PDFS}")
    exit()

print(f"Dosto Copión: Leyendo {len(archivos)} PDFs de la carpeta de ejemplos...")

with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as txt_salida:
    for archivo in archivos:
        ruta_completa = os.path.join(CARPETA_PDFS, archivo)
        txt_salida.write(f"\n{'='*60}\n")
        txt_salida.write(f"--- ARCHIVO: {archivo} ---\n")
        txt_salida.write(f"{'='*60}\n\n")
        
        try:
            with pdfplumber.open(ruta_completa) as pdf:
                for i, pagina in enumerate(pdf.pages, start=1):
                    texto = pagina.extract_text()
                    txt_salida.write(f"--- PÁGINA {i} ---\n")
                    if texto:
                        txt_salida.write(texto + "\n")
                    else:
                        txt_salida.write("[No se pudo extraer texto de esta página o es una imagen]\n")
        except Exception as e:
            txt_salida.write(f"[ERROR AL PROCESAR EL ARCHIVO]: {e}\n")
            
        txt_salida.write("\n\n")

print(f"¡Listo, Gustavo! Se generó el archivo en:\n-> {ARCHIVO_SALIDA}")