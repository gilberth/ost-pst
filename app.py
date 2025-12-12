#!/usr/bin/env python3
"""
OST to PST Converter - Aplicación con interfaz gráfica
Convierte archivos OST (Offline Storage Table) a PST (Personal Storage Table)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import threading
from pathlib import Path
from converter import OSTtoPSTConverter, ConversionMethod

class OSTtoPSTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OST to PST Converter")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        self.converter = OSTtoPSTConverter()
        self.conversion_running = False

        self.setup_ui()
        self.check_dependencies()

    def setup_ui(self):
        """Configura la interfaz de usuario"""
        # Estilo
        style = ttk.Style()
        style.theme_use('clam')

        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configurar expansión
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

        # Título
        title_label = ttk.Label(main_frame, text="Convertidor OST a PST",
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)

        # Sección archivo de entrada
        input_frame = ttk.LabelFrame(main_frame, text="Archivo OST de Entrada", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        input_frame.columnconfigure(1, weight=1)

        ttk.Label(input_frame, text="Archivo:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.input_path = tk.StringVar()
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=50)
        self.input_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)

        self.browse_input_btn = ttk.Button(input_frame, text="Examinar...",
                                          command=self.browse_input)
        self.browse_input_btn.grid(row=0, column=2, padx=5)

        # Sección archivo de salida
        output_frame = ttk.LabelFrame(main_frame, text="Archivo PST de Salida", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(1, weight=1)

        ttk.Label(output_frame, text="Archivo:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        self.output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)

        self.browse_output_btn = ttk.Button(output_frame, text="Examinar...",
                                           command=self.browse_output)
        self.browse_output_btn.grid(row=0, column=2, padx=5)

        # Opciones de conversión
        options_frame = ttk.LabelFrame(main_frame, text="Opciones de Conversión", padding="10")
        options_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(options_frame, text="Método:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.method = tk.StringVar(value="hybrid_mbox")
        methods = [
            ("Híbrido MBOX (Recomendado)", "hybrid_mbox"),
            ("Aspose.Email (Comercial)", "aspose"),
            ("Win32COM (Windows+Outlook)", "win32com"),
            ("readpst (Solo extracción)", "readpst"),
            ("libpff (Solo lectura)", "libpff")
        ]

        # Distribuir en dos filas para mejor visualización
        for i, (text, value) in enumerate(methods):
            row = i // 3
            col = (i % 3) + 1
            ttk.Radiobutton(options_frame, text=text, variable=self.method,
                          value=value).grid(row=row, column=col, padx=5, pady=2, sticky=tk.W)

        # Botones de acción
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=15)

        self.convert_btn = ttk.Button(button_frame, text="Convertir",
                                     command=self.start_conversion,
                                     style='Accent.TButton')
        self.convert_btn.pack(side=tk.LEFT, padx=5)

        self.cancel_btn = ttk.Button(button_frame, text="Cancelar",
                                    command=self.cancel_conversion,
                                    state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="Salir",
                  command=self.root.quit).pack(side=tk.LEFT, padx=5)

        # Área de log/progreso
        log_frame = ttk.LabelFrame(main_frame, text="Registro de Conversión", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=70,
                                                 wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # Barra de estado
        self.status_var = tk.StringVar(value="Listo")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var,
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E))

    def browse_input(self):
        """Abre diálogo para seleccionar archivo OST"""
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo OST",
            filetypes=[
                ("Archivos OST", "*.ost"),
                ("Todos los archivos", "*.*")
            ]
        )
        if filename:
            self.input_path.set(filename)
            # Sugerir nombre de salida
            if not self.output_path.get():
                output = str(Path(filename).with_suffix('.pst'))
                self.output_path.set(output)

    def browse_output(self):
        """Abre diálogo para seleccionar ubicación de salida PST"""
        filename = filedialog.asksaveasfilename(
            title="Guardar archivo PST como",
            defaultextension=".pst",
            filetypes=[
                ("Archivos PST", "*.pst"),
                ("Todos los archivos", "*.*")
            ]
        )
        if filename:
            self.output_path.set(filename)

    def log(self, message, level="INFO"):
        """Agrega mensaje al área de log"""
        self.log_text.config(state=tk.NORMAL)
        timestamp = Path(__file__).stem  # Placeholder para timestamp simple
        self.log_text.insert(tk.END, f"[{level}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def check_dependencies(self):
        """Verifica que las dependencias necesarias estén instaladas"""
        self.log("Verificando dependencias del sistema...", "INFO")

        available = self.converter.check_available_methods()

        if available['readpst']:
            self.log("✓ readpst encontrado", "INFO")
        else:
            self.log("✗ readpst no encontrado (instalar con: apt-get install pst-utils)", "WARNING")

        if available['hybrid_mbox']:
            self.log("✓ Método híbrido MBOX disponible", "INFO")
        else:
            self.log("✗ Método híbrido MBOX no disponible (requiere readpst)", "WARNING")

        if available['libpff']:
            self.log("✓ libpff encontrado", "INFO")
        else:
            self.log("✗ libpff no encontrado (instalar con: pip install pypff)", "WARNING")

        if available['aspose']:
            self.log("✓ Aspose.Email encontrado", "INFO")
        else:
            self.log("✗ Aspose.Email no encontrado (comercial: pip install Aspose.Email-for-Python-via-NET)", "WARNING")

        if available['win32com']:
            self.log("✓ win32com encontrado (Windows + Outlook)", "INFO")
        else:
            self.log("✗ win32com no disponible (solo Windows: pip install pywin32)", "WARNING")

        if not any(available.values()):
            self.log("⚠ ADVERTENCIA: No se encontraron herramientas de conversión instaladas", "ERROR")
            messagebox.showwarning(
                "Dependencias faltantes",
                "No se encontraron herramientas de conversión instaladas.\n\n"
                "Por favor instale al menos una de las siguientes:\n"
                "- readpst (pst-utils) - RECOMENDADO\n"
                "- pypff (Python libpff)\n"
                "- Aspose.Email-for-Python-via-NET (comercial)\n"
                "- pywin32 (solo Windows con Outlook)"
            )

    def validate_inputs(self):
        """Valida que los campos estén correctamente llenados"""
        input_file = self.input_path.get()
        output_file = self.output_path.get()

        if not input_file:
            messagebox.showerror("Error", "Por favor seleccione un archivo OST de entrada")
            return False

        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"El archivo de entrada no existe:\n{input_file}")
            return False

        if not output_file:
            messagebox.showerror("Error", "Por favor especifique un archivo PST de salida")
            return False

        # Verificar extensión
        if not input_file.lower().endswith('.ost'):
            response = messagebox.askyesno(
                "Advertencia",
                f"El archivo de entrada no tiene extensión .ost\n¿Continuar de todos modos?"
            )
            if not response:
                return False

        # Verificar si el archivo de salida existe
        if os.path.exists(output_file):
            response = messagebox.askyesno(
                "Confirmar sobrescritura",
                f"El archivo de salida ya existe:\n{output_file}\n\n¿Desea sobrescribirlo?"
            )
            if not response:
                return False

        return True

    def start_conversion(self):
        """Inicia el proceso de conversión en un hilo separado"""
        if not self.validate_inputs():
            return

        self.conversion_running = True
        self.convert_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.browse_input_btn.config(state=tk.DISABLED)
        self.browse_output_btn.config(state=tk.DISABLED)

        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        self.progress.start()
        self.status_var.set("Convirtiendo...")

        # Ejecutar conversión en hilo separado
        thread = threading.Thread(target=self.run_conversion)
        thread.daemon = True
        thread.start()

    def run_conversion(self):
        """Ejecuta la conversión (llamado desde hilo separado)"""
        try:
            input_file = self.input_path.get()
            output_file = self.output_path.get()
            method_str = self.method.get()

            # Mapear string a enum
            method_map = {
                'readpst': ConversionMethod.READPST,
                'libpff': ConversionMethod.LIBPFF,
                'aspose': ConversionMethod.ASPOSE,
                'win32com': ConversionMethod.WIN32COM,
                'hybrid_mbox': ConversionMethod.HYBRID_MBOX
            }
            method = method_map.get(method_str, ConversionMethod.HYBRID_MBOX)

            self.log(f"Iniciando conversión de {os.path.basename(input_file)}", "INFO")
            self.log(f"Método: {method_str}", "INFO")

            # Ejecutar conversión
            success, message = self.converter.convert(
                input_file,
                output_file,
                method,
                progress_callback=self.log
            )

            # Actualizar UI en el hilo principal
            self.root.after(0, self.conversion_complete, success, message)

        except Exception as e:
            self.root.after(0, self.conversion_complete, False, str(e))

    def conversion_complete(self, success, message):
        """Llamado cuando la conversión se completa"""
        self.progress.stop()
        self.conversion_running = False

        self.convert_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.browse_input_btn.config(state=tk.NORMAL)
        self.browse_output_btn.config(state=tk.NORMAL)

        if success:
            self.log(f"✓ Conversión completada exitosamente", "SUCCESS")
            self.log(f"Archivo guardado en: {self.output_path.get()}", "SUCCESS")
            self.status_var.set("Conversión completada")
            messagebox.showinfo("Éxito", "La conversión se completó exitosamente")
        else:
            self.log(f"✗ Error en la conversión: {message}", "ERROR")
            self.status_var.set("Error en conversión")
            messagebox.showerror("Error", f"Error durante la conversión:\n\n{message}")

    def cancel_conversion(self):
        """Cancela la conversión en progreso"""
        if messagebox.askyesno("Confirmar", "¿Está seguro de que desea cancelar la conversión?"):
            self.conversion_running = False
            self.log("Conversión cancelada por el usuario", "WARNING")
            self.status_var.set("Cancelado")
            # Note: Implementación real requeriría comunicación con el proceso de conversión


def main():
    root = tk.Tk()
    app = OSTtoPSTApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
