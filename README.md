# Dosto 📚🤖

**Dosto** es una aplicación de escritorio moderna desarrollada en Python para automatizar la extracción de datos de facturas de diferentes editoriales (PDF) y convertirlas en reportes unificados de Excel listos para la carga en sistemas de gestión.

El sistema incluye un **motor multieditorial** capaz de interpretar estructuras variables de texto y un **sistema dinámico de equivalencias** para mapear códigos internos a sus respectivos códigos de barras o ISBN reales.

---

## 🚀 Características principales

* 📊 **Detector Automático:** Identifica al proveedor analizando el encabezado (Soporta SBS, Planeta, Kapelusz e Ivrea...POR AHORA!!).
* 🧼 **Limpieza de Datos:** Elimina guiones de los ISBNs, separa títulos pegados y normaliza cantidades decimales.
* 💸 **Cálculo de Descuentos:** Procesa descuentos específicos por fila o tasas globales al pie de la factura según el proveedor.
* 🗺️ **Mapeo Inteligente:** Lee un archivo `equivalencias.xlsx` local para reemplazar de forma automática los códigos internos de los artículos.
* 🖥️ **Interfaz Gráfica (GUI):** Diseñada con `CustomTkinter`, con soporte nativo para modo oscuro y navegación nativa de archivos.

---

## 🛠️ Instalación y Requisitos (Desarrollo)

Si necesitás correr o editar el entorno de desarrollo, asegurate de tener **Python 3.13+** e instalar las siguientes dependencias:

```bash
pip install pdfplumber pandas openpyxl customtkinter