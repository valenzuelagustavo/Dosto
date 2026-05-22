import pdfplumber

# Poné el nombre de la factura que te dio la alerta
archivo_pdf = "4-20931 13042026 Losada.pdf" 

print(f"Abriendo {archivo_pdf} para analizar su estructura...\n")

with pdfplumber.open(archivo_pdf) as pdf:
    # Leemos solo la primera página para ver el encabezado y el inicio de la tabla
    primera_pagina = pdf.pages[0]
    texto = primera_pagina.extract_text()
    
    print("--- COPIÁ DESDE ACÁ ABAJO ---")
    print(texto)
    print("--- HASTA ACÁ ARRIBA ---")