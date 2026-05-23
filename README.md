# Dosto - Administrador de Documentos Comerciales (v1.5)

**Dosto** es una herramienta de automatización interna desarrollada en Python para la gestión y auditoría de documentos comerciales (Facturas, Notas de Crédito y Remitos) en **Papillon Libros S.A.** El sistema elimina la carga manual de datos pesados, procesando extractos masivos de texto de PDFs y realizando controles cruzados automáticos de stock y devoluciones.

---

## 🚀 Funcionalidades Principales

La aplicación cuenta con una interfaz gráfica moderna dividida en dos módulos clave:

### 1. Extractor Individual
* **Conversión Inteligente:** Transforma facturas y notas de crédito en formato PDF a plantillas limpias de Excel (`.xlsx`).
* **Mapeo Automático:** Vincula e inyecta códigos de barra, ISBNs de 13 dígitos y porcentajes de descuento fila por fila.
* **Soporte Multi-Editorial:** Motor optimizado nativamente con expresiones regulares (RegEx) y huellas digitales (CUIT/Metadatos) para grandes distribuidoras:
  * **Penguin Random House** (Soporta etiquetas FSC `#`, variaciones de alineación y formatos de descuento con guion `40-`).
  * **Santillana**
  * **Maipue** (Identificación robusta por CUIT `27-10970866-1` y gestión de ajustes globales sin ítems).
  * **Heliasta** (Extracción de descuentos globales en el pie de página e inmunidad a encabezados basados en imágenes).

### 2. Comparador de Devoluciones (¡Nuevo en v1.5!)
* **Consolidación Multi-Origen:** Permite cargar **múltiples Notas de Crédito (PDF)** y **múltiples Remitos del sistema (Excel)** en simultáneo.
* **Trazabilidad por Caja/Remito:** Consolida las unidades enviadas horizontalmente y detalla en qué remito físico salió cada libro.
* **Mapeo de Títulos por Posición Física:** Asocia de forma infalible cada ISBN con su nombre de libro real, aislando de forma segura las celdas de texto o cuadernos sin código de barra.
* **Redactor de Reclamos Automático:** Genera un archivo de texto independiente (`RECLAMO_MAIL_AUTOMATICO.txt`) con una carta formal alineada en columnas, lista para copiar y pegar en un mail a la editorial en caso de faltantes.

---

## 📊 Estructura del Reporte de Control Cruzado

El comparador genera un informe quirúrgico detallando el estado de cada artículo devuelto:

| ISBN | Título | Cant_remito_caja1 | Cant_remito_caja2 | Total_Remitos | Cant_NC | Diferencia | Estado |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :--- |
| 9789500757584 | ROSAURA A LAS DIEZ | 1 | 2 | 3 | 3 | 0 | OK |
| 9788410096042 | LA ALEGRIA DE DECIR NO | 5 | 0 | 5 | 4 | -1 | FALTAN 1 u. |

---

## 🛠️ Tecnologías Utilizadas

* **Python 3.13**
* **CustomTkinter:** Interfaz gráfica de usuario con soporte nativo para Modo Oscuro/Claro del sistema.
* **Pandas:** Motor de análisis de datos para la consolidación, agrupamiento y cruce de inventarios (`Outer Joins`).
* **Pdfplumber & Re:** Extracción y parseo de texto crudo mediante expresiones regulares avanzadas.

---

## 📦 Instalación y Uso en la Oficina

El sistema está preparado para correr en entorno local sin necesidad de usar la terminal en el día a día.

1. Clonar el repositorio en el directorio de trabajo:
   ```bash
   git clone [https://github.com/](https://github.com/)[tu-usuario]/Dosto.git

2. Asegurarse de tener las dependencias instaladas:
    ´´bash
    pip install customtkinter pandas pdfplumber openpyxl

3. Ejecutar el asistente de Windows haciendo doble click en el lanzador:
    iniciar_dosto.bat