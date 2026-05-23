import os
import re
import pandas as pd
import customtkinter as ctk
from extractor import procesar_factura_pdf

def comparar_devolucion_masiva(lista_rutas_nc, lista_rutas_remitos, ruta_equivalencias="equivalencias.xlsx"):
    """
    Compara múltiples NC (PDF) contra múltiples Remitos (Excel).
    Consolida datos, incluye nombres de libros extraídos del PDF y
    genera un reporte Excel + un archivo .txt con el mail de reclamo automatizado.
    """
    try:
        # --- 1. PROCESAR Y CONSOLIDAR NOTAS DE CRÉDITO + MAPEO DE TÍTULOS ---
        print(f"[Comparador] Procesando {len(lista_rutas_nc)} Notas de Crédito...")
        df_nc_consolidado = pd.DataFrame(columns=['ISBN', 'Cant_NC'])
        mapeo_titulos = {} # Diccionario para guardar {ISBN: "Título del Libro"}
        archivos_nc_temporales = []

        # Para deducir el nombre de la editorial en el encabezado del mail
        editorial_detectada = "la editorial"

        for ruta_nc in lista_rutas_nc:
            exito, msg = procesar_factura_pdf(ruta_nc, ruta_equivalencias)
            if not exito:
                return False, f"Error al procesar el PDF {os.path.basename(ruta_nc)}: {msg}"
            
            nombre_base = os.path.splitext(os.path.basename(ruta_nc))[0]
            ruta_tmp = os.path.join(os.path.dirname(ruta_nc), f"procesado_{nombre_base}.xlsx")
            archivos_nc_temporales.append(ruta_tmp)

            # Cargar datos temporales de esta NC
            df_tmp = pd.read_excel(ruta_tmp)
            df_tmp['ISBN'] = df_tmp['ISBN'].astype(str).str.strip()
            
            # Intentar deducir la editorial del mensaje de éxito si es posible
            if "PLANETA" in msg.upper(): editorial_detectada = "Planeta"
            elif "PENGUIN" in msg.upper(): editorial_detectada = "Penguin Random House"
            elif "SANTILLANA" in msg.upper(): editorial_detectada = "Santillana"
            elif "HELIASTA" in msg.upper(): editorial_detectada = "Heliasta"

            # Recolectar títulos vinculando el ISBN (si la columna existe en el procesado intermedio)
            col_titulo_tmp = None
            for c in df_tmp.columns:
                c_upper = c.upper()
                if ('TITULO' in c_upper or 'DESC' in c_upper) and 'DESCUENTO' not in c_upper and 'DESC.%' not in c_upper:
                    col_titulo_tmp = c
                    break

            if col_titulo_tmp:
                for _, fila in df_tmp.iterrows():
                    isbn_f = str(fila['ISBN']).strip()
                    desc_f = str(fila[col_titulo_tmp]).strip()
                    # Si el título es un número puro de dos dígitos (como el 40 del descuento), lo ignoramos
                    if isbn_f and desc_f and desc_f.upper() != "NAN" and not (desc_f.isdigit() and len(desc_f) <= 2):
                        mapeo_titulos[isbn_f] = desc_f

            df_tmp_agrupado = df_tmp.groupby('ISBN', as_index=False)['Cantidad'].sum()
            df_tmp_agrupado.rename(columns={'Cantidad': 'Cant_NC'}, inplace=True)
            df_nc_consolidado = pd.concat([df_nc_consolidado, df_tmp_agrupado], ignore_index=True)

        # Consolidación final de todas las NCs juntas
        df_nc_consolidado = df_nc_consolidado.groupby('ISBN', as_index=False)['Cant_NC'].sum()

        # --- 2. PROCESAR Y CONSOLIDAR TODOS LOS REMITOS (EXCEL) ---
        print(f"[Comparador] Procesando {len(lista_rutas_remitos)} remitos del sistema...")
        df_remitos_consolidados = pd.DataFrame(columns=['ISBN'])
        columnas_cantidades_remitos = []
        
        for ruta_remito in lista_rutas_remitos:
            nombre_corto_remito = os.path.splitext(os.path.basename(ruta_remito))[0]
            col_dinamica_remito = f"Cant_{nombre_corto_remito}"
            columnas_cantidades_remitos.append(col_dinamica_remito)
            
            df_r = pd.read_excel(ruta_remito)
            df_r.columns = df_r.columns.str.strip().str.upper()
            
            col_isbn = 'ISBN' if 'ISBN' in df_r.columns else ('CODIGO' if 'CODIGO' in df_r.columns else df_r.columns[0])
            col_cant = 'CANTIDAD' if 'CANTIDAD' in df_r.columns else ('CANT' if 'CANT' in df_r.columns else df_r.columns[1])
            
            df_r_limpio = df_r[[col_isbn, col_cant]].copy()
            df_r_limpio.columns = ['ISBN', col_dinamica_remito]
            df_r_limpio['ISBN'] = df_r_limpio['ISBN'].astype(str).str.strip()
            
            # Recolectar títulos del remito si el PDF no los tenía (ej. artículos genéricos o Maipue)
            col_desc_rem = next((c for c in df_r.columns if 'DESC' in c or 'TITULO' in c or 'NOMBRE' in c), None)
            if col_desc_rem:
                for _, fila in df_r.iterrows():
                    isbn_f = str(fila[col_isbn]).strip()
                    desc_f = str(fila[col_desc_rem]).strip()
                    if isbn_f and desc_f and isbn_f not in mapeo_titulos:
                        mapeo_titulos[isbn_f] = desc_f

            df_r_agrupado = df_r_limpio.groupby('ISBN', as_index=False)[col_dinamica_remito].sum()
            
            if df_remitos_consolidados.empty:
                df_remitos_consolidados = df_r_agrupado
            else:
                df_remitos_consolidados = pd.merge(df_remitos_consolidados, df_r_agrupado, on='ISBN', how='outer')
        
        df_remitos_consolidados.fillna(0, inplace=True)
        df_remitos_consolidados['Total_Remitos'] = df_remitos_consolidados[columnas_cantidades_remitos].sum(axis=1)
        
        # --- 3. CRUCE FINAL Y DISCREPANCIAS ---
        print("[Comparador] Cruzando datos y calculando diferencias...")
        df_final = pd.merge(df_remitos_consolidados, df_nc_consolidado, on='ISBN', how='outer').fillna(0)
        
        # Convertir a enteros de forma segura (blindaje anti-cuadernos de colorear)
        columnas_a_entero = columnas_cantidades_remitos + ['Total_Remitos', 'Cant_NC']
        for col in columnas_a_entero:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0).astype(int)
            
        df_final['ISBN'] = df_final['ISBN'].astype(str)
        df_final['Diferencia'] = df_final['Cant_NC'] - df_final['Total_Remitos']
        
        def determinar_estado(row):
            if row['Diferencia'] == 0: return "OK"
            return f"FALTAN {abs(row['Diferencia'])} u." if row['Diferencia'] < 0 else f"SOBRAN {row['Diferencia']} u."
                
        df_final['Estado'] = df_final.apply(determinar_estado, axis=1)
        
        # Inyectar columna de Títulos basada en nuestro mapa recolectado
        df_final['Título'] = df_final['ISBN'].map(lambda x: mapeo_titulos.get(x, "Descripción no encontrada en origen"))
        
        # Reordenar columnas estéticamente incluyendo el Título al principio
        columnas_ordenadas = ['ISBN', 'Título'] + columnas_cantidades_remitos + ['Total_Remitos', 'Cant_NC', 'Diferencia', 'Estado']
        df_final = df_final[columnas_ordenadas]
        
        # --- 4. EXPORTAR EXCEL Y GENERAR EL RECLAMO .TXT ---
        carpeta_destino = os.path.dirname(lista_rutas_nc[0])
        ruta_excel = os.path.join(carpeta_destino, "reporte_control_devolucion_CONSOLIDADO.xlsx")
        ruta_txt = os.path.join(carpeta_destino, "RECLAMO_MAIL_AUTOMATICO.txt")
        
        df_final.to_excel(ruta_excel, index=False)
        
        # Filtrar solo los ítems donde nos reconocieron menos unidades (Diferencia negativa)
        df_faltantes = df_final[df_final['Diferencia'] < 0]
        
        # Redacción automática y humana del archivo .txt
        with open(ruta_txt, "w", encoding="utf-8") as f:
            f.write(f"Asunto: Reclamo por unidades faltantes en Nota de Crédito - Papillon Libros\n\n")
            f.write(f"Estimados representantes de {editorial_detectada},\n\n")
            f.write("Junto con saludar, nos comunicamos con ustedes para solicitar la revisión del procesamiento ")
            f.write("de nuestra última devolución de material. Al realizar el control cruzado automatizado ")
            f.write("entre nuestros remitos internos de salida y las Notas de Crédito recibidas, ")
            f.write("hemos detectado discrepancias en las cantidades reconocidas.\n\n")
            f.write("A continuación, detallamos los títulos que presentan diferencias (unidades faltantes):\n")
            f.write("-" * 85 + "\n")
            
            if df_faltantes.empty:
                f.write("No se detectaron unidades faltantes. ¡El control cruzado dio todo OK!\n")
            else:
                for _, fila in df_faltantes.iterrows():
                    titulo_corto = str(fila['Título'])[:45]
                    f.write(f"• ISBN: {fila['ISBN']:<14} | {titulo_corto:<45} | Enviado: {fila['Total_Remitos']:>2} u. / Reconocido: {fila['Cant_NC']:>2} u. -> (Faltan {abs(fila['Diferencia'])} u.)\n")
            
            f.write("-" * 85 + "\n\n")
            f.write("Agradecemos desde ya su gestión para verificar estos movimientos en sus registros ")
            f.write("y quedamos a la espera de la correspondiente nota de crédito complementaria por el saldo afectado.\n\n")
            f.write("Quedamos a su entera disposición ante cualquier consulta.\n\n")
            f.write("Saludos cordiales,\n")
            f.write("Dpto. de Administración y Recepción\n")
            f.write("Papillon Libros S.A.\n")

        # Limpieza de archivos basura temporales
        for path_tmp in archivos_nc_temporales:
            if os.path.exists(path_tmp): os.remove(path_tmp)
            
        return True, f"¡Control completado!\n\nSe generaron en la carpeta:\n1. 📊 reporte_control_devolucion_CONSOLIDADO.xlsx\n2. 📄 RECLAMO_MAIL_AUTOMATICO.txt"
        
    except Exception as e:
        return False, f"Error en el comparador masivo: {e}"