import os
import customtkinter as ctk
from customtkinter import filedialog
# Importamos la función que creamos en el otro archivo
from extractor import procesar_factura_pdf

# Configuración estética general de CustomTkinter
ctk.set_appearance_mode("System")  # Detecta si Windows está en modo oscuro o claro
ctk.set_default_color_theme("blue") # Tema de color general (botones, acentos)

class AppDosto(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración de la ventana principal
        self.title("Dosto - Extractor de Facturas")
        self.geometry("550x450")
        self.resizable(False, False)

        self.ruta_pdf_seleccionado = ""

        # --- DISEÑO DE LA INTERFAZ ---
        
        # 1. Título principal
        self.label_titulo = ctk.CTkLabel(self, text="DOSTO", font=ctk.CTkFont(size=28, weight="bold"))
        self.label_titulo.pack(pady=(20, 5))
        
        self.label_sub = ctk.CTkLabel(self, text="Automatización de Facturas de Editoriales", font=ctk.CTkFont(size=12))
        self.label_sub.pack(pady=(0, 20))

        # 2. Caja de estado del Excel de equivalencias
        estado_eq = "Detectado" if os.path.exists("equivalencias.xlsx") else "No encontrado (Se procesará sin mapeo)"
        color_eq = "green" if os.path.exists("equivalencias.xlsx") else "orange"
        self.label_eq = ctk.CTkLabel(self, text=f"Archivo de Equivalencias: {estado_eq}", text_color=color_eq, font=ctk.CTkFont(size=11, weight="bold"))
        self.label_eq.pack(pady=5)

       # 3. Botón para seleccionar el PDF (Corregido 'padx')
        self.btn_seleccionar = ctk.CTkButton(self, text="Seleccionar Factura PDF", command=self.abrir_selector_archivos, height=40, font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_seleccionar.pack(pady=15, fill="x", padx=40)

        # 4. Etiqueta para mostrar el archivo elegido
        self.label_archivo = ctk.CTkLabel(self, text="Ningún archivo seleccionado", text_color="gray", wraplength=450)
        self.label_archivo.pack(pady=5)

        # 5. Botón principal de Procesar (Corregido 'padx')
        self.btn_procesar = ctk.CTkButton(self, text="PROCESAR FACTURA", command=self.ejecutar_procesamiento, state="disabled", fg_color="gray", height=45, font=ctk.CTkFont(size=15, weight="bold"))
        self.btn_procesar.pack(pady=15, fill="x", padx=40)

        # 6. Consola de estado / Registro de acciones (Corregido 'padx')
        self.txt_consola = ctk.CTkTextbox(self, height=120, activate_scrollbars=True, font=ctk.CTkFont(family="Consolas", size=11))
        self.txt_consola.pack(pady=(10, 20), fill="x", padx=40)
        
    # --- FUNCIONES DE LA INTERFAZ ---

    def abrir_selector_archivos(self):
        # Abre la ventana clásica de Windows para elegir archivos, filtrando solo PDFs
        archivo = filedialog.askopenfilename(
            title="Seleccionar factura PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        
        if archivo:
            self.ruta_pdf_seleccionado = archivo
            nombre_corto = os.path.basename(archivo)
            
            # Actualizamos la interfaz
            self.label_archivo.configure(text=f"Archivo listo: {nombre_corto}", text_color="white" if ctk.get_appearance_mode() == "Dark" else "black")
            self.btn_procesar.configure(state="normal", fg_color=("#1f538d", "#1f76b4")) # Restauramos color azul original
            self.imprimir_consola(f"Archivo cargado correctamente:\n{archivo}")

    def ejecutar_procesamiento(self):
        if not self.ruta_pdf_seleccionado:
            return

        self.imprimir_consola("\nIniciando extracción de datos...")
        self.btn_procesar.configure(state="disabled") # Evita doble clicks durante el proceso
        
        # Llamamos al motor que está en extractor.py
        exito, mensaje = procesar_factura_pdf(self.ruta_pdf_seleccionado)
        
        # Actualizamos la consola visual con el resultado del motor
        self.imprimir_consola(mensaje)
        
        if exito:
            # Ponemos el botón en verde si todo salió bien
            self.btn_procesar.configure(fg_color="green", text="¡COMPLETADO CON ÉXITO!")
        else:
            # Ponemos el botón en rojo si hubo algún error
            self.btn_procesar.configure(fg_color="crimson", text="ERROR EN EL PROCESO")
            
        # Reestablecemos el botón después de 3 segundos para poder procesar otro archivo
        self.after(3000, self.restaurar_boton_procesar)

    def restaurar_boton_procesar(self):
        if self.ruta_pdf_seleccionado:
            self.btn_procesar.configure(state="normal", fg_color=("#1f538d", "#1f76b4"), text="PROCESAR FACTURA")
        else:
            self.btn_procesar.configure(state="disabled", fg_color="gray", text="PROCESAR FACTURA")

    def imprimir_consola(self, texto):
        # Función auxiliar para escribir líneas en nuestra mini terminal de pantalla
        self.txt_consola.insert("end", texto + "\n")
        self.txt_consola.see("end") # Auto-scroll hacia abajo

# Ejecución de la aplicación
if __name__ == "__main__":
    app = AppDosto()
    app.mainloop()