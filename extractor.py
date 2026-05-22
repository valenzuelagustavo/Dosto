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
    descuento_global_factura = 0

    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            # 2. DETECTOR DE EDITORIAL (Analiza la primera página)
            primera_pág = pdf.pages[0].extract_text()
            texto_identificacion = primera_pág.upper() if primera_pág else ""
            
            # --- SISTEMA SELECTOR DE LÓGICA MULTIEDITORIAL ---
            if "STRATFORD" in texto_identificacion:
                editorial = "SBS"
                cabecera_tabla = "Caja Cant. Código Detalle"
                patron = r"^\d+\s+(\d+)\s+(\S+).*? \$\s+([\d.,]+)\s+(\d+)\s+\$\s+[\d.,]+$"
                
            elif "PLANETA" in texto_identificacion:
                editorial = "PLANETA"
                cabecera_tabla = "ARTICULO CANTIDAD P. UNIT IMPORTE"
                # RegEx Planeta Blindada: ISBN (1), salta título y opcionales códigos entre paréntesis, captura Cantidad (2), captura Precio (3)
                patron = r"^([\d-]+).*? (?:$$\d+$$ )?(\d+)\s+([\d.,]+)\s+[\d.,]+$"
                
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
                    descuento_global_factura = int(float(match_desc.group(1).replace(',', '.')))
                    
            elif "ILHSA" in texto_identificacion or "ATENEO" in texto_identificacion:
                editorial = "ILHSA"
                cabecera_tabla = "ISBN TITULO CANTIDAD PRECIO UNITARIO ALICUOTA IVA % DTO DESCUENTO IMPORTE NETO"
                patron = r"^(\d{13}).*?\s+(\d+)\s+([\d.,]+)\s+\d+\s+(\d+)"
                
            elif "EDITORIALGUADAL" in texto_identificacion.replace(" ", "") or "CÓDIGO ISBN DESCRIPCIÓN CANT." in texto_identificacion:
                editorial = "GUADAL"
                cabecera_tabla = "Código ISBN Descripción Cant."
                patron = r"^(\d+)\s+(\d{13})\s+.*?\s+(\d+)\s+([\d.,]+)\s+-?([\d.,]+)\s*%"
                
            elif "URANO" in texto_identificacion:
                editorial = "URANO"
                cabecera_tabla = "Ped. ("
                patron = r"^(\d{13}).*?\s+([\d.,]+)\s+([\d.,]+)\s*%\s+[\d.,]+"
                
            elif "CODIGO EJEMPLARES TITULO PRECIO DTO NETO" in texto_identificacion:
                editorial = "CONTINENTE"
                cabecera_tabla = "Codigo Ejemplares Titulo Precio Dto Neto"
                patron = r"^(\S+)\s+(\d+)\s+.*?\s+([\d.,]+)\s+(\d+)\s+[\d.,]+$"
                
            elif "LONGSELLER" in texto_identificacion:
                editorial = "LONGSELLER"
                cabecera_tabla = "CANT DESCRIPCIÓN DTO PVP IMPORTE"
                patron = r"^(\d+)\s+(\d+)\s+.*?\s+([\d.,]+)\s+([\d.,]+)\s+[\d.,]+$"
                
            elif "IMAGINADOR" in texto_identificacion:
                editorial = "IMAGINADOR"
                cabecera_tabla = "CODIGO T¡TULO SEUDONIMO"
                patron = r"^(\d+)\s+.*?\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+[\d.,]+$"
                
            elif "KEL EDICIONES" in texto_identificacion or "C.KEL ISBN DESCRIPCIÓN" in texto_identificacion:
                editorial = "KEL"
                cabecera_tabla = "C.Kel ISBN Descripción"
                patron = r"^\d+\s+(\d{13})\s+.*?\s+([\d.,]+)\s+([\d.,]+)\s+[\d.,]+\s+(\d+)\s+[\d.,]+"
                
            elif "RED DEL LIBRO" in texto_identificacion or "SATORI" in texto_identificacion or "IT CANT EAN DESCRIPCIÓN" in texto_identificacion:
                editorial = "SISTEMA_RED_SATORI"
                cabecera_tabla = "IT Cant EAN Descripción"
                patron = r"^\d+\s+(\d+)\s+(\d{13})\s+.*?\s+([\d.,]+)\s+(\d+[\d.,]*)\s*%\s+[\d.,]+"
                
            elif "MANDIOCA" in texto_identificacion:
                editorial = "MANDIOCA"
                cabecera_tabla = "Condición de Venta:"
                patron = r"^(\d+)\s+(\s*[\d-]{10,18}\s*)\s+.*?\s+([\d.,]+)\s+([\d.,]+)\s+[\d.,]+$"
                
            elif "FONDO DE CULTURA" in texto_identificacion or "FCE" in texto_identificacion or "ISBN ARTICULO AUTOR CANT." in texto_identificacion:
                editorial = "FCE"
                cabecera_tabla = "ISBN ARTICULO AUTOR CANT. PRECIO %DESCT IMPORTE"
                # RegEx FCE: ISBN (1), salta texto, Cantidad decimal (2), Precio Unitario (3), % Descuento decimal (4)
                patron = r"^(\d{13})\s+.*?\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+[\d.,]+"
                
            elif "EDELVIVES" in texto_identificacion:
                editorial = "EDELVIVES"
                cabecera_tabla = "Cód. ISBN Descripción Código artículo Cantidad"
                # RegEx Edelvives: ISBN (1), salta texto/código, Cantidad (2), salta "Unidades", Precio Unitario (3), % Descuento (4)
                patron = r"^(\d{13})\s+.*?\s+(\d+)\s+Unidades\s+([\d.,]+)\s+(\d+)\s+[\d.,]+"
                
            elif "LOSADA" in texto_identificacion or "ISBN TITULO AUTOR CANT." in texto_identificacion:
                editorial = "LOSADA"
                cabecera_tabla = "ISBN TITULO AUTOR Cant. P.Unit. Des.% Importe"
                # RegEx Losada: ISBN (1), salta texto, Cantidad (2), Precio Unitario (3), Descuento (4)
                patron = r"^(\d{13})\s+.*?\s+(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+[\d.,]+$"
                
            elif "GRUPAL" in texto_identificacion:
                editorial = "GRUPAL"
                cabecera_tabla = "Cant ISBN Descripción Despacho Precio Desc Total"
                patron = r"^(\d+)\s+(\S+).*?\s+\S+\s+\$\s+([\d.,]+)\s+([\d.,]+)\s+\$\s+[\d.,]+"
                
            elif "ATLANTIDA" in texto_identificacion:
                editorial = "ATLANTIDA"
                cabecera_tabla = "Código Descripción Cant. Código Barras Precio unit. Desc. Precio Neto"
                patron = r"^\d+\s+.*?\s+(\d+)\s+(\d{13})\s+([\d.,]+)\s+-?([\d.,]+)"
                
            elif "AZ EDITORA" in texto_identificacion or "NUEVO EXTREMO" in texto_identificacion or "EDITORIAL KIER" in texto_identificacion or "EUDEBA" in texto_identificacion or "PROMETEO" in texto_identificacion or "OVNI PRESS" in texto_identificacion or "M4 EDITORIAL" in texto_identificacion:
                editorial = "SISTEMA_ESTANDAR_B"
                cabecera_tabla = "Cant ISBN Descripción"
                patron = r"^(\d+)\s+(\S+).*?\$\s+([\d.,]+)\s+([\d.,]+)\s+\$\s+[\d.,]+"
                
            elif "V&R EDITORAS" in texto_identificacion:
                editorial = "VYR"
                cabecera_tabla = "Cant ISBN Descripción Precio Desc Neto Total"
                patron = r"^(\d+)\s+(\d{13}).*?\$\s+([\d.,]+)\s+([\d.,]+)\s+\$\s+[\d.,]+"
                
            elif "VESTALES" in texto_identificacion:
                editorial = "VESTALES"
                cabecera_tabla = "CCaannttiiddaadd IISSBBNN"
                patron = r"^(\d+)\s+(\d{13})\s+.*?\$\s+([\d.,]+)\s+([\d.,]+)\s*%\s+\$\s+[\d.,]+"
                
            elif "MANOLITO BOOKS" in texto_identificacion:
                editorial = "MANOLITO"
                cabecera_tabla = "Código Concepto Colección Cantidad"
                patron = r"^(\d{13})\s+.*?\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+[\d.,]+"
                
            elif "HELIASTA" in texto_identificacion or "37250" in texto_identificacion:
                editorial = "HELIASTA"
                cabecera_tabla = "PLAZO DE PAGO"
                patron = r"^(\d+)\s+(\d{13})\s+.*?\s+([\d.,]+)\s+([\d.,]+)$"
                match_desc = re.search(r"Desc\.\s+([\d.,]+)\s*%", texto_identificacion)
                if match_desc:
                    descuento_global_factura = int(float(match_desc.group(1).replace(',', '.')))

            elif "SANTILLANA" in texto_identificacion:
                editorial = "SANTILLANA"
                cabecera_tabla = "Artículo Cant. Descripción Unitario Bruto %dto Subtotal"
                # RegEx: Código interno (1), Cantidad (2), salta texto, Precio Unitario (3), % Descuento (4)
                patron = r"^(\d+)\s+(\d+)\s+.*?([\d.,]+)\s+[\d.,]+\s+([\d.,]+)\s+[\d.,]+"

            elif "MAIPUE" in texto_identificacion or "10970866" in texto_identificacion:
                editorial = "MAIPUE"
                cabecera_tabla = "Descripción Precio Desc Total" # Cabecera estándar del sistema
                # RegEx Maipue: ISBN (1), salta título, Precio Lista (2), % Descuento (3), salta el Neto
                patron = r"^(\d{13})\s+.*?\s+([\d.,]+)\s+([\d.,]+)\s+\$\s+[\d.,]+$"

            else:
                return False, "Editorial no reconocida en el encabezado del PDF."

            print(f"[Extractor] Procesando con motor: {editorial}")

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
                        # Ignoramos líneas de transporte intermedia que rompen Guadal en multifactura
                        if editorial == "GUADAL" and "Transporte :" in linea:
                            continue
                            
                        # Frenos de corte de seguridad
                        if editorial in ["SISTEMA_ESTANDAR_B", "VYR", "SISTEMA_RED_SATORI", "FCE"] and ("Lineas:" in linea or "Total:" in linea or "Cantidad de Ejemplares" in linea or "SubTotal Neto" in linea or "SUBTOTAL ARS" in linea):
                            break
                        if editorial == "PLANETA" and ("Subtotal Factura" in linea or "TOTAL EJEMPLARES" in linea):
                            break
                        if editorial == "EDELVIVES" and "Subtotal" in linea:
                            break
                        if editorial == "LOSADA" and ("Vendedor:" in linea or "NETO GRAVADO" in linea):
                            break
                        if editorial == "IMAGINADOR" and ("Vendedor:" in linea or "NETO GRAVADO" in linea):
                            break
                        if editorial == "MANDIOCA" and "Subtotal :" in linea:
                            break
                        if editorial == "LONGSELLER" and ("Son:" in linea or "Unidades:" in linea):
                            break
                        if editorial == "ATLANTIDA" and "Subtotal" in linea:
                            break
                        if editorial == "GRUPAL" and "Lineas:" in linea:
                            break
                        if editorial == "VESTALES" and ("Ejemplares:" in linea or "Total:" in linea):
                            break
                        if editorial == "MANOLITO" and "Importe Neto" in linea:
                            break
                        if editorial == "HELIASTA" and "Subtotal" in linea:
                            break
                        if editorial == "MAIPUE" and ("Lineas:" in linea or "Total:" in linea):
                            break
                        if "TOTAL $" in linea or "TOTAL EJEMPLARES" in linea or "SON:" in linea:
                            break
                        if editorial == "SANTILLANA" and "Total Bruto" in linea:
                            break
                            
                        resultado = re.search(patron, linea)
                        
                        if resultado:
                            # --- ASIGNACIÓN DE VARIABLES SEGÚN CADA MOTOR ---
                            if editorial in ["SISTEMA_ESTANDAR_B", "VYR", "GRUPAL", "SISTEMA_RED_SATORI"]:
                                cantidad = int(resultado.group(1))
                                codigo_detectado = str(resultado.group(2)).strip().replace("-", "")
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                                
                            elif editorial in ["FCE", "MANOLITO"]:
                                codigo_detectado = str(resultado.group(1)).strip().replace("-", "")
                                cant_sucia = resultado.group(2).replace('.', '').replace(',', '.')
                                cantidad = int(float(cant_sucia))
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                                
                            elif editorial == "EDELVIVES":
                                codigo_detectado = str(resultado.group(1)).strip().replace("-", "")
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                                
                            elif editorial == "LOSADA":
                                codigo_detectado = str(resultado.group(1)).strip().replace("-", "")
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                                
                            elif editorial == "IMAGINADOR":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                                
                            elif editorial == "MANDIOCA":
                                cantidad = int(resultado.group(1))
                                codigo_detectado = str(resultado.group(2)).strip().replace("-", "").replace(" ", "")
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                                
                            elif editorial == "LONGSELLER":
                                cantidad = int(resultado.group(1))
                                codigo_detectado = str(resultado.group(2)).strip()
                                descuento = int(float(resultado.group(3).replace(',', '.')))
                                precio_unitario = resultado.group(4)
                                
                            elif editorial == "ATLANTIDA":
                                cantidad = int(resultado.group(1))
                                codigo_detectado = str(resultado.group(2)).strip()
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                                
                            elif editorial == "VESTALES":
                                cantidad = int(resultado.group(1))
                                codigo_detectado = str(resultado.group(2)).strip()
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                                
                            elif editorial == "HELIASTA":
                                cantidad = int(resultado.group(1))
                                codigo_detectado = str(resultado.group(2)).strip()
                                precio_unitario = resultado.group(3)
                                descuento = descuento_global_factura
                                
                            elif editorial == "SBS":
                                cantidad = int(resultado.group(1))
                                codigo_detectado = str(resultado.group(2)).strip()
                                precio_unitario = resultado.group(3)
                                descuento = int(resultado.group(4))
                            
                            elif editorial == "PLANETA":
                                codigo_sucio = resultado.group(1)
                                codigo_detectado = str(codigo_sucio.replace("-", "")).strip()
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = 0  # Descuento global al pie
                                
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
                                descuento = descuento_global_factura
                                
                            elif editorial == "ILHSA":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = int(resultado.group(4))
                                
                            elif editorial == "GUADAL":
                                codigo_detectado = str(resultado.group(2)).strip()
                                cantidad = int(resultado.group(3))
                                precio_unitario = resultado.group(4)
                                descuento = int(float(resultado.group(5).replace(',', '.')))
                                
                            elif editorial == "URANO":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad = 1
                                precio_unitario = resultado.group(2)
                                descuento = int(float(resultado.group(3).replace(',', '.')))
                                
                            elif editorial == "CONTINENTE":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = int(resultado.group(4))
                            
                            elif editorial == "SANTILLANA":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad = int(resultado.group(2))
                                precio_unitario = resultado.group(3)
                                descuento = int(float(resultado.group(4).replace(',', '.')))
                            
                            elif editorial == "MAIPUE":
                                codigo_detectado = str(resultado.group(1)).strip()
                                cantidad = 1 # Las líneas de este formato liquidan por unidad en la tabla
                                precio_unitario = resultado.group(2)
                                descuento = int(float(resultado.group(3).replace(',', '.')))

                            # --- CRUCE POR DICCIONARIO DE EQUIVALENCIAS ---
                            if codigo_detectado in base_equivalencias:
                                isbn_final = base_equivalencias[codigo_detectado]
                            else:
                                isbn_final = codigo_detectado
                            
                            precio_limpio = float(precio_unitario.replace('.', '').replace(',', '.'))
                            
                            # Si es Planeta, el descuento final real del pie se inyecta luego en bloque, 
                            # pero seteamos 0 por fila como pide el formato original.
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