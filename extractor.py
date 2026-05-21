import re
import os
import pdfplumber
import pandas as pd

def procesar_factura_pdf(ruta_pdf, ruta_equivalencias="equivalencias.xlsx"):
    base_equivalencias = {}

    # 1. CARGAR BASE DE DATOS DESDE EXCEL
    if os.path.exists(ruta_equivalencias):
        try:
            df_eq = pd.read_excel(ruta_equivalencias)
            df_eq.columns = df_eq.columns.str.strip().str.lower()
            df_eq['codigo_editorial'] = df_eq['codigo_editorial'].astype(str).str.strip()
            df_eq['isbn_real'] = df_eq['isbn_real'].astype(str).str.strip()
            base_equivalencias = dict(zip(df_eq['codigo_editorial'], df_eq['isbn_real']))
        except Exception as e:
            return False, f"Error al leer el Excel de equivalencias: {e}"

    datos_articulos = []
    descuento_global_ivrea = 0

    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            # 2. DETECTOR DE EDITORIAL
            primera_pág = pdf.pages[0].extract_text()
            texto_identificacion = primera_pág.upper() if primera_pág else ""
            
            if "STRATFORD" in texto_identificacion:
                editorial = "SBS"
                cabecera_tabla = "Caja Cant. Código Detalle"
                patron = r"^\d+\s+(\d+)\s+(\S+).*? \$\s+([\d.,]+)\s+(\d+)\s+\$\s+[\d.,]+$"
                
            elif "PLANETA" in texto_identificacion:
                editorial = "PLANETA"
                cabecera_tabla = "ARTICULO CANTIDAD P. UNIT IMPORTE"
                patron = r"^([\d-]+).*?\s+(\d+)\s+([\d.,]+)\s+[\d.,]+$"
                
            elif "KAPELUSZ" in texto_identificacion:
                editorial = "KAPELUSZ"
                cabecera_tabla = "Artículo Cant. Descripción Remito Unitario Bruto %dto Subtotal"
                patron = r"^(\d+)\s+(\d+)\s+.*?21\d{8}\s+([\d.,]+)\s+(?:[\d.,]+\s+)([\d.,]+)"
                
            elif "IVREA" in texto_identificacion or "SIBIU" in texto_identificacion:
                editorial = "IVREA"
                cabecera_tabla = "ARTÍCULO DESCRIPCIÓN COD. BARRAS CANTIDAD PRECIO UNIT. IMPORTE"
                patron = r"^\S+\s+.*?\s+(\d{13})\s+([\d.,]+)\s+\$\s+([\d.,]+)"
                
                match_desc = re.search(r"DESCUENTO:\s+-?([\d.,]+)\s*%", texto_identificacion)
                if match_desc:
                    descuento_global_ivrea = int(float(match_desc.group(1).replace(',', '.')))
                    
            elif "ILHSA" in texto_identificacion or "ATENEO" in texto_identificacion:
                editorial = "ILHSA"
                cabecera_tabla = "ISBN TITULO CANTIDAD PRECIO UNITARIO ALICUOTA IVA % DTO DESCUENTO IMPORTE NETO"
                # RegEx para Ilhsa:
                # ^(\d{13}) -> Captura los 13 dígitos del ISBN (1) que vienen pegados al título
                # .*?\s+(\d+) -> Salta el título y captura la Cantidad (2)
                # \s+([\d.,]+) -> Captura el Precio Unitario (3)
                # \s+\d+\s+(\d+) -> Salta la alícuota de IVA y captura el % Descuento (4)
                patron = r"^(\d{13}).*?\s+(\d+)\s+([\d.,]+)\s+\d+\s+(\d+)"
                
            else:
                return False, "Editorial no reconocida en el encabezado del PDF."

            print(f"[Extractor] Detectada editorial: {editorial}")

            # 3. PROCESAMIENTO DE LAS PÁGINAS
            for numero_pagina, pagina in enumerate(pdf.pages, start=1):
                texto_completo = pagina.extract_text()
                if not texto_completo:
                    continue
                    
                lineas = texto_completo.split("\n")
                empezar_a_leer_libros = True if numero_pagina > 1 else False
                
                for linea in lineas:
                    linea = linea.strip()
                    
                    if cabecera_tabla in linea:
                        empezar_a_leer_libros = True
                        continue
                    
                    if empezar_a_leer_libros:
                        # Condiciones de corte según editorial
                        if editorial == "IVREA" and ("SUB. TOTAL:" in linea or "Total Libros:" in linea):
                            break
                        if editorial == "ILHSA" and ("REGIMEN DE TRANSPARENCIA" in linea or "SUBTOTAL $:" in linea):
                            break
                        if "Total Bruto" in linea or "TOTAL $" in linea or "TOTAL EJEMPLARES" in linea or "SON:" in linea:
                            break
                            
                        resultado = re.search(patron, linea)
                        
                        if resultado:
                            if editorial == "SBS":
                                cantidad = int(resultado.group(1))
                                codigo_detectado = str(resultado.group(2)).strip()
                                precio_unitario = resultado.group(3)
                                descuento = int(resultado.group(4))
                            
                            elif editorial == "PLANETA":
                                codigo_sucio = resultado.group(1)
                                codigo_detectado = str(codigo_sucio.replace("-", "")).strip()
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = 0
                                
                            elif editorial == "KAPELUSZ":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento_sucio = resultado.group(4)
                                descuento = int(float(descuento_sucio.replace(',', '.')))
                                
                            elif editorial == "IVREA":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad_sucia = resultado.group(2).replace('.', '').replace(',', '.')
                                cantidad = int(float(cantidad_sucia))
                                precio_unitario = resultado.group(3)
                                descuento = descuento_global_ivrea
                                
                            elif editorial == "ILHSA":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = int(resultado.group(4))
                            
                            # Mapeo de equivalencias
                            if codigo_detectado in base_equivalencias:
                                isbn_final = base_equivalencias[codigo_detectado]
                            else:
                                isbn_final = codigo_detectado
                            
                            precio_limpio = float(precio_unitario.replace('.', '').replace(',', '.'))
                            
                            datos_articulos.append({
                                "ISBN": isbn_final,
                                "Cantidad": cantidad,
                                "Precio": precio_limpio,
                                "Descuento": descuento
                            })

        # 4. EXPORTACIÓN DEL EXCEL FINAL
        if datos_articulos:
            df = pd.DataFrame(datos_articulos)
            df['ISBN'] = df['ISBN'].astype(str)
            
            nombre_base = os.path.splitext(os.path.basename(ruta_pdf))[0]
            nombre_excel = f"procesado_{nombre_base}.xlsx"
            
            ruta_salida_excel = os.path.join(os.path.dirname(ruta_pdf), nombre_excel)
            df.to_excel(ruta_salida_excel, index=False)
            
            return True, f"¡Éxito! Se procesó la factura de {editorial}.\nGenerado: {nombre_excel}\nTotal de artículos: {len(datos_articulos)}"
        else:
            return False, "No se pudieron extraer datos válidos de las tablas del PDF."

    except Exception as e:
        return False, f"Ocurrió un error inesperado al procesar: {e}"