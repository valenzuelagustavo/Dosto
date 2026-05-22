import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from extractor import procesar_factura_pdf
from comparador import comparar_devolucion_multiple

# Configuración estética base de CustomTkinter
ctk.set_appearance_mode("System")  # Adopta el modo del sistema (Oscuro/Claro)
ctk.set_default_color_theme("blue")

class AppDosto(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Dosto - Administrador de Documentos Comerciales v1.5")
        self.geometry("680x580")
        self.resizable(False, False)

        # Variables para el Modo Extractor Individual
        self.ruta_pdf_individual = ""

        # Variables para el Modo Comparador Múltiple
        self.ruta_pdf_nc = ""
        self.lista_rutas_remitos = []

        # --- CREACIÓN DE PESTAÑAS (TABVIEW) ---
        self.tabview = ctk.CTkTabview(self, width=640, height=540)
        self.tabview.pack(padx=20, pady=20)
        
        self.tabview.add("Extractor Individual")
        self.tabview.add("Comparador de Devoluciones")

        # Configurar los paneles individuales
        self.setup_tab_extractor()
        self.setup_tab_comparador()

    # =========================================================================
    # 🛠️ PANEL 1: EXTRACTOR INDIVIDUAL (Lógica clásica)
    # =========================================================================
    def setup_tab_extractor(self):
        tab = self.tabview.tab("Extractor Individual")

        lbl_titulo = ctk.CTkLabel(tab, text="Extractor Automático de Facturas", font=ctk.CTkFont(size=18, weight="bold"))
        lbl_titulo.pack(pady=15)

        lbl_desc = ctk.CTkLabel(tab, text="Seleccioná un PDF de cualquier editorial para convertirlo en un Excel limpio\nmapeando automáticamente los códigos por tu base de equivalencias.", font=ctk.CTkFont(size=12))
        lbl_desc.pack(pady=5)

        self.btn_cargar_ind = ctk.CTkButton(tab, text="Seleccionar Factura / NC (PDF)", command=self.cargar_pdf_individual, height=40)
        self.btn_cargar_ind.pack(pady=20)

        self.lbl_status_ind = ctk.CTkLabel(tab, text="Ningún archivo seleccionado", text_color="gray", wraplength=550)
        self.lbl_status_ind.pack(pady=10)

        self.btn_procesar_ind = ctk.CTkButton(tab, text="Procesar Documento", command=self.ejecutar_extractor_individual, state="disabled", fg_color="green", hover_color="darkgreen", height=45, font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_procesar_ind.pack(pady=30)

    def cargar_pdf_individual(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        if ruta:
            self.ruta_pdf_individual = ruta
            self.lbl_status_ind.configure(text=f"Archivo cargado:\n{os.path.basename(ruta)}", text_color="white")
            self.btn_procesar_ind.configure(state="normal")

    def ejecutar_extractor_individual(self):
        self.btn_procesar_ind.configure(state="disabled", text="Procesando...")
        self.update_idletasks()
        
        exito, msg = procesar_factura_pdf(self.ruta_pdf_individual)
        
        if exito:
            messagebox.showinfo("¡Éxito!", msg)
        else:
            messagebox.showerror("Error", msg)
            
        self.btn_procesar_ind.configure(state="normal", text="Procesar Documento")

    # =========================================================================
    # 📊 PANEL 2: COMPARADOR DE DEVOLUCIONES (Lógica nueva)
    # =========================================================================
    def setup_tab_comparador(self):
        tab = self.tabview.tab("Comparador de Devoluciones")

        lbl_titulo = ctk.CTkLabel(tab, text="Control Cruzado de Devoluciones", font=ctk.CTkFont(size=18, weight="bold"))
        lbl_titulo.pack(pady=10)

        lbl_desc = ctk.CTkLabel(tab, text="Cruza una Nota de Crédito contra uno o varios remitos del sistema.\nConsolida cantidades del mismo libro y detecta faltantes o sobrantes.", font=ctk.CTkFont(size=12))
        lbl_desc.pack(pady=5)

        # --- SECTOR NOTA DE CRÉDITO ---
        frame_nc = ctk.CTkFrame(tab)
        frame_nc.pack(fill="x", padx=20, pady=10)

        self.btn_cargar_nc = ctk.CTkButton(frame_nc, text="1. Cargar NC (PDF)", command=self.cargar_pdf_nc, width=150)
        self.btn_cargar_nc.pack(side="left", padx=10, pady=10)

        self.lbl_status_nc = ctk.CTkLabel(frame_nc, text="No se cargó la Nota de Crédito", text_color="gray", anchor="w")
        self.lbl_status_nc.pack(side="left", fill="x", expand=True, padx=10)

        # --- SECTOR REMITOS MÚLTIPLES ---
        frame_remitos = ctk.CTkFrame(tab)
        frame_remitos.pack(fill="both", expand=True, padx=20, pady=5)

        # Subframe para los botones de acción de remitos
        frame_botones_r = ctk.CTkFrame(frame_remitos, fg_color="transparent")
        frame_botones_r.pack(side="left", padx=10, pady=10, fill="y")

        self.btn_sumar_remito = ctk.CTkButton(frame_botones_r, text="2. Añadir Remito (Excel)", command=self.sumar_remito_lista, width=160, fg_color="#1f538d")
        self.btn_sumar_remito.pack(pady=5)

        self.btn_limpiar_remitos = ctk.CTkButton(frame_botones_r, text="Limpiar Cola", command=self.limpiar_cola_remitos, width=160, fg_color="#a83232", hover_color="#822424")
        self.btn_limpiar_remitos.pack(pady=5)

        # Caja de texto para simular la lista de archivos cargados
        self.txt_lista_remitos = ctk.CTkTextbox(frame_remitos, height=100, font=ctk.CTkFont(size=11))
        self.txt_lista_remitos.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.txt_lista_remitos.insert("0.0", "Cola de remitos vacía...\nAgrega uno o más archivos .xlsx")
        self.txt_lista_remitos.configure(state="disabled")

        # --- BOTÓN DE EJECUCIÓN FINAL ---
        self.btn_comparar = ctk.CTkButton(tab, text="Ejecutar Control de Discrepancias", command=self.ejecutar_comparador_masivo, state="disabled", fg_color="green", hover_color="darkgreen", height=45, font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_comparar.pack(pady=15)

    def cargar_pdf_nc(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        if ruta:
            self.ruta_pdf_nc = ruta
            self.lbl_status_nc.configure(text=f"NC cargada: {os.path.basename(ruta)}", text_color="white")
            self.verificar_listo_para_comparar()

    def sumar_remito_lista(self):
        rutas = filedialog.askopenfilenames(filetypes=[("Archivos Excel", "*.xlsx")])
        if rutas:
            for ruta in rutas:
                if ruta not in self.lista_rutas_remitos:
                    self.lista_rutas_remitos.append(ruta)
            self.actualizar_vista_lista_remitos()
            self.verificar_listo_para_comparar()

    def limpiar_cola_remitos(self):
        self.lista_rutas_remitos = []
        self.actualizar_vista_lista_remitos()
        self.verificar_listo_para_comparar()

    def actualizar_vista_lista_remitos(self):
        self.txt_lista_remitos.configure(state="normal")
        self.txt_lista_remitos.delete("0.0", "end")
        
        if not self.lista_rutas_remitos:
            self.txt_lista_remitos.insert("0.0", "Cola de remitos vacía...\nAgrega uno o más archivos .xlsx")
        else:
            for idx, ruta in enumerate(self.lista_rutas_remitos, start=1):
                self.txt_lista_remitos.insert("end", f"[{idx}] {os.path.basename(ruta)}\n")
                
        self.txt_lista_remitos.configure(state="disabled")

    def verificar_listo_para_comparar(self):
        # Para poder ejecutar el control necesitamos obligatoriamente la NC y al menos un remito
        if self.ruta_pdf_nc and self.lista_rutas_remitos:
            self.btn_comparar.configure(state="normal")
        else:
            self.btn_comparar.configure(state="disabled")

    def ejecutar_comparador_masivo(self):
        self.btn_comparar.configure(state="disabled", text="Comparando y consolidando...")
        self.update_idletasks()
        
        exito, msg = comparar_devolucion_multiple(self.ruta_pdf_nc, self.lista_rutas_remitos)
        
        if exito:
            messagebox.showinfo("¡Control Completado!", msg)
            self.limpiar_cola_remitos()  # Limpia la cola tras un proceso exitoso
            self.ruta_pdf_nc = ""
            self.lbl_status_nc.configure(text="No se cargó la Nota de Crédito", text_color="gray")
        else:
            messagebox.showerror("Error en el Proceso", msg)
            
        self.btn_comparar.configure(state="normal", text="Ejecutar Control de Discrepancias")

if __name__ == "__main__":
    app = AppDosto()
    app.mainloop()