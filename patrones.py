import pdfplumber

# Poné el nombre de la factura que te dio la alerta
archivo_pdf = "FacturaC 0013-00052156_07-05-2026 13-05-06 Ivrea.pdf" 

print(f"Abriendo {archivo_pdf} para analizar su estructura...\n")

with pdfplumber.open(archivo_pdf) as pdf:
    # Leemos solo la primera página para ver el encabezado y el inicio de la tabla
    primera_pagina = pdf.pages[0]
    texto = primera_pagina.extract_text()
    
    print("--- COPIÁ DESDE ACÁ ABAJO ---")
    print(texto)
    print("--- HASTA ACÁ ARRIBA ---")