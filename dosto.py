import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from extractor import procesar_factura_pdf
from comparador import comparar_devolucion_masiva

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

        # Variables para el Modo Comparador Múltiple (Estructura de listas gigantes)
        self.lista_rutas_nc = []
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
    # 📊 PANEL 2: COMPARADOR DE DEVOLUCIONES (Soporte Masivo Multi-Origen)
    # =========================================================================
    def setup_tab_comparador(self):
        tab = self.tabview.tab("Comparador de Devoluciones")

        lbl_titulo = ctk.CTkLabel(tab, text="Control Cruzado de Devoluciones", font=ctk.CTkFont(size=18, weight="bold"))
        lbl_titulo.pack(pady=5)

        # --- CONTENEDOR DE DOS COLUMNAS SIMÉTRICAS ---
        frame_listas = ctk.CTkFrame(tab, fg_color="transparent")
        frame_listas.pack(fill="both", expand=True, padx=10, pady=5)

        # COLUMNA IZQUIERDA: NOTAS DE CRÉDITO (PDF)
        frame_col_nc = ctk.CTkFrame(frame_listas)
        frame_col_nc.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        btn_add_nc = ctk.CTkButton(frame_col_nc, text="1. Añadir NCs (PDF)", command=self.sumar_nc_lista, fg_color="#1f538d")
        btn_add_nc.pack(pady=10, padx=10, fill="x")

        self.txt_lista_nc = ctk.CTkTextbox(frame_col_nc, height=120, font=ctk.CTkFont(size=11))
        self.txt_lista_nc.pack(fill="both", expand=True, padx=10, pady=5)
        self.txt_lista_nc.insert("0.0", "Ninguna NC cargada...\nAgrega archivos .pdf")
        self.txt_lista_nc.configure(state="disabled")

        btn_clear_nc = ctk.CTkButton(frame_col_nc, text="Limpiar NCs", command=self.limpiar_cola_nc, fg_color="#a83232", hover_color="#822424")
        btn_clear_nc.pack(pady=5, padx=10, fill="x")

        # COLUMNA DERECHA: REMITOS INTERNOS (EXCEL)
        frame_col_rem = ctk.CTkFrame(frame_listas)
        frame_col_rem.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        btn_add_rem = ctk.CTkButton(frame_col_rem, text="2. Añadir Remitos (Excel)", command=self.sumar_remito_lista, fg_color="#1f538d")
        btn_add_rem.pack(pady=10, padx=10, fill="x")

        self.txt_lista_remitos = ctk.CTkTextbox(frame_col_rem, height=120, font=ctk.CTkFont(size=11))
        self.txt_lista_remitos.pack(fill="both", expand=True, padx=10, pady=5)
        self.txt_lista_remitos.insert("0.0", "Ningún remito cargado...\nAgrega archivos .xlsx")
        self.txt_lista_remitos.configure(state="disabled")

        btn_clear_rem = ctk.CTkButton(frame_col_rem, text="Limpiar Remitos", command=self.limpiar_cola_remitos, fg_color="#a83232", hover_color="#822424")
        btn_clear_rem.pack(pady=5, padx=10, fill="x")

        # --- BOTÓN DE EJECUCIÓN FINAL ---
        self.btn_comparar = ctk.CTkButton(tab, text="Ejecutar Control de Discrepancias Masivo", command=self.ejecutar_comparador_masivo, state="disabled", fg_color="green", hover_color="darkgreen", height=45, font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_comparar.pack(pady=15)

    # =========================================================================
    # 🎛️ MÉTODOS DE SOPORTE INTERNOS (Alineados dentro de la clase AppDosto)
    # =========================================================================
    def sumar_nc_lista(self):
        rutas = filedialog.askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
        if rutas:
            for ruta in rutas:
                if ruta not in self.lista_rutas_nc:
                    self.lista_rutas_nc.append(ruta)
            self.actualizar_vista_lista_nc()
            self.verificar_listo_para_comparar()

    def limpiar_cola_nc(self):
        self.lista_rutas_nc = []
        self.actualizar_vista_lista_nc()
        self.verificar_listo_para_comparar()

    def actualizar_vista_lista_nc(self):
        self.txt_lista_nc.configure(state="normal")
        self.txt_lista_nc.delete("0.0", "end")
        if not self.lista_rutas_nc:
            self.txt_lista_nc.insert("0.0", "Ninguna NC cargada...\nAgrega archivos .pdf")
        else:
            for idx, ruta in enumerate(self.lista_rutas_nc, start=1):
                self.txt_lista_nc.insert("end", f"[{idx}] {os.path.basename(ruta)}\n")
        self.txt_lista_nc.configure(state="disabled")

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
            self.txt_lista_remitos.insert("0.0", "Ningún remito cargado...\nAgrega archivos .xlsx")
        else:
            for idx, ruta in enumerate(self.lista_rutas_remitos, start=1):
                self.txt_lista_remitos.insert("end", f"[{idx}] {os.path.basename(ruta)}\n")
        self.txt_lista_remitos.configure(state="disabled")

    def verificar_listo_para_comparar(self):
        # Valida que la cola tenga por lo menos una NC y un Remito cargado para activarse
        if self.lista_rutas_nc and self.lista_rutas_remitos:
            self.btn_comparar.configure(state="normal")
        else:
            self.btn_comparar.configure(state="disabled")

    def ejecutar_comparador_masivo(self):
        self.btn_comparar.configure(state="disabled", text="Consolidando universos...")
        self.update_idletasks()
        
        exito, msg = comparar_devolucion_masiva(self.lista_rutas_nc, self.lista_rutas_remitos)
        
        if exito:
            messagebox.showinfo("¡Control Completado!", msg)
            self.limpiar_cola_remitos()
            self.limpiar_cola_nc()
        else:
            messagebox.showerror("Error en el Proceso", msg)
            self.btn_comparar.configure(state="normal", text="Ejecutar Control de Discrepancias Masivo")

# --- ARRANQUE NATIVO DEL PROGRAMA ---
if __name__ == "__main__":
    app = AppDosto()
    app.mainloop()