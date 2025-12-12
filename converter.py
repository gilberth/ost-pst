"""
Módulo de conversión OST a PST
Soporta múltiples métodos de conversión
"""

import os
import subprocess
import shutil
from enum import Enum
from pathlib import Path
from typing import Callable, Tuple, Optional


class ConversionMethod(Enum):
    """Métodos de conversión disponibles"""
    READPST = "readpst"
    LIBPFF = "libpff"
    ASPOSE = "aspose"


class OSTtoPSTConverter:
    """Clase principal para conversión de archivos OST a PST"""

    def __init__(self):
        self.available_methods = self.check_available_methods()

    def check_available_methods(self) -> dict:
        """Verifica qué métodos de conversión están disponibles"""
        available = {
            'readpst': False,
            'libpff': False,
            'aspose': False
        }

        # Verificar readpst
        if shutil.which('readpst'):
            available['readpst'] = True

        # Verificar libpff (pypff)
        try:
            import pypff
            available['libpff'] = True
        except ImportError:
            pass

        # Verificar Aspose.Email
        try:
            import aspose.email
            available['aspose'] = True
        except ImportError:
            pass

        return available

    def convert(self,
                input_file: str,
                output_file: str,
                method: ConversionMethod = ConversionMethod.READPST,
                progress_callback: Optional[Callable] = None) -> Tuple[bool, str]:
        """
        Convierte un archivo OST a PST

        Args:
            input_file: Ruta al archivo OST de entrada
            output_file: Ruta al archivo PST de salida
            method: Método de conversión a utilizar
            progress_callback: Función opcional para reportar progreso

        Returns:
            Tuple[bool, str]: (éxito, mensaje)
        """

        def log(message: str, level: str = "INFO"):
            if progress_callback:
                progress_callback(message, level)

        # Validar entrada
        if not os.path.exists(input_file):
            return False, f"El archivo de entrada no existe: {input_file}"

        # Verificar método disponible
        method_key = method.value
        if not self.available_methods.get(method_key, False):
            return False, f"El método '{method_key}' no está disponible. Por favor instale las dependencias necesarias."

        # Ejecutar conversión según el método
        try:
            if method == ConversionMethod.READPST:
                return self._convert_with_readpst(input_file, output_file, log)
            elif method == ConversionMethod.LIBPFF:
                return self._convert_with_libpff(input_file, output_file, log)
            elif method == ConversionMethod.ASPOSE:
                return self._convert_with_aspose(input_file, output_file, log)
            else:
                return False, f"Método de conversión no soportado: {method}"

        except Exception as e:
            return False, f"Error durante la conversión: {str(e)}"

    def _convert_with_readpst(self,
                             input_file: str,
                             output_file: str,
                             log: Callable) -> Tuple[bool, str]:
        """
        Convierte usando readpst (de pst-utils)

        Nota: readpst extrae el contenido de archivos OST/PST pero no crea
        directamente un archivo PST. Esta implementación extrae y luego
        usa otras herramientas para reconstruir.
        """
        log("Usando método readpst", "INFO")

        try:
            # Crear directorio temporal para extracción
            temp_dir = Path(output_file).parent / f".temp_extraction_{Path(input_file).stem}"
            temp_dir.mkdir(exist_ok=True)

            log(f"Extrayendo contenido a directorio temporal: {temp_dir}", "INFO")

            # Ejecutar readpst para extraer
            cmd = [
                'readpst',
                '-o', str(temp_dir),  # Directorio de salida
                '-e',                  # Export mode (mbox)
                '-D',                  # Include deleted items
                input_file
            ]

            log(f"Ejecutando: {' '.join(cmd)}", "INFO")

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=3600  # 1 hora timeout
            )

            if result.returncode != 0:
                log(f"Error en readpst: {result.stderr}", "ERROR")
                return False, f"readpst falló: {result.stderr}"

            log("Extracción completada", "INFO")
            log(f"Salida: {result.stdout}", "INFO")

            # Nota: readpst no crea archivos PST directamente
            # Para crear un PST real, necesitaríamos:
            # 1. Usar Microsoft Outlook API (solo Windows)
            # 2. Usar Aspose.Email u otra biblioteca comercial
            # 3. Convertir los archivos mbox/eml extraídos

            # Por ahora, intentamos usar readpst-pst si está disponible
            # o crear una estructura compatible

            log("NOTA: readpst extrae el contenido pero no crea PST nativo", "WARNING")
            log("Intentando conversión alternativa...", "INFO")

            # Intentar crear PST usando método alternativo
            return self._create_pst_from_extraction(temp_dir, output_file, log)

        except subprocess.TimeoutExpired:
            return False, "La conversión excedió el tiempo límite de 1 hora"
        except Exception as e:
            return False, f"Error en conversión con readpst: {str(e)}"

    def _convert_with_libpff(self,
                            input_file: str,
                            output_file: str,
                            log: Callable) -> Tuple[bool, str]:
        """Convierte usando libpff (pypff)"""
        log("Usando método libpff (pypff)", "INFO")

        try:
            import pypff

            log(f"Abriendo archivo OST: {input_file}", "INFO")

            # Abrir archivo OST
            ost_file = pypff.file()
            ost_file.open(input_file)

            log(f"Archivo OST abierto. Carpetas raíz: {ost_file.number_of_root_folders}", "INFO")

            # pypff puede leer pero no escribir PST directamente
            # Necesitamos extraer y reconstruir

            log("NOTA: pypff puede leer OST/PST pero no escribir PST nativo", "WARNING")
            log("Extrayendo datos para conversión...", "INFO")

            # Extraer información
            root = ost_file.get_root_folder()

            # Aquí necesitaríamos implementar la lógica completa de extracción
            # y reconstrucción del PST

            ost_file.close()

            return False, "pypff no soporta escritura directa de PST. Use Aspose.Email o conversión manual."

        except ImportError:
            return False, "pypff no está instalado. Instale con: pip install pypff"
        except Exception as e:
            return False, f"Error con libpff: {str(e)}"

    def _convert_with_aspose(self,
                           input_file: str,
                           output_file: str,
                           log: Callable) -> Tuple[bool, str]:
        """Convierte usando Aspose.Email (método comercial más robusto)"""
        log("Usando método Aspose.Email", "INFO")

        try:
            import aspose.email as ae
            from aspose.email.storage.pst import PersonalStorage, FileFormat

            log(f"Abriendo archivo OST: {input_file}", "INFO")

            # Abrir archivo OST
            ost = PersonalStorage.from_file(input_file)

            log(f"Archivo OST abierto. Formato: {ost.format}", "INFO")
            log(f"Total de items: {ost.store.total_items_count if hasattr(ost.store, 'total_items_count') else 'N/A'}", "INFO")

            log(f"Convirtiendo a PST: {output_file}", "INFO")

            # Convertir a PST
            # Método simple: copiar todas las carpetas y contenido
            pst = PersonalStorage.create(output_file, FileFormat.UNICODE)

            # Copiar carpetas recursivamente
            self._copy_folders_recursive(ost.root_folder, pst.root_folder, log)

            # Cerrar archivos
            pst.dispose()
            ost.dispose()

            log("Conversión completada exitosamente", "SUCCESS")

            return True, "Conversión completada exitosamente usando Aspose.Email"

        except ImportError:
            return False, ("Aspose.Email no está instalado.\n"
                          "Instale con: pip install Aspose.Email-for-Python-via-NET\n"
                          "Nota: Esta es una biblioteca comercial que requiere licencia.")
        except Exception as e:
            return False, f"Error con Aspose.Email: {str(e)}"

    def _copy_folders_recursive(self, source_folder, dest_folder, log: Callable):
        """Copia carpetas recursivamente de OST a PST (para Aspose)"""
        try:
            import aspose.email as ae

            # Copiar subcarpetas
            for subfolder in source_folder.get_sub_folders():
                log(f"Copiando carpeta: {subfolder.display_name}", "INFO")

                # Crear carpeta en destino
                new_folder = dest_folder.add_sub_folder(subfolder.display_name)

                # Copiar mensajes
                messages = subfolder.get_contents()
                for msg_info in messages:
                    msg = subfolder.get_message(msg_info.entry_id_string)
                    new_folder.add_message(msg)

                # Recursión para subcarpetas
                self._copy_folders_recursive(subfolder, new_folder, log)

        except Exception as e:
            log(f"Error copiando carpeta: {str(e)}", "ERROR")

    def _create_pst_from_extraction(self,
                                   temp_dir: Path,
                                   output_file: str,
                                   log: Callable) -> Tuple[bool, str]:
        """
        Intenta crear un PST desde archivos extraídos
        Esta es una implementación simplificada
        """
        log("Intentando crear PST desde extracción...", "INFO")

        # Verificar si hay archivos extraídos
        extracted_files = list(temp_dir.glob('**/*'))

        if not extracted_files:
            return False, "No se extrajeron archivos del OST"

        log(f"Se extrajeron {len(extracted_files)} archivos/carpetas", "INFO")

        # Para una conversión real, necesitaríamos:
        # 1. Parsear los archivos mbox/eml extraídos
        # 2. Usar una biblioteca que pueda escribir PST (como Aspose)
        # 3. Reconstruir la estructura de carpetas

        # Por ahora, informamos al usuario
        msg = (
            f"Se extrajo el contenido a: {temp_dir}\n"
            f"Para crear un PST real, use el método Aspose.Email\n"
            f"o importe estos archivos en Microsoft Outlook manualmente."
        )

        log(msg, "WARNING")

        return False, msg


def main():
    """Función principal para uso en línea de comandos"""
    import argparse

    parser = argparse.ArgumentParser(description='Convertir archivos OST a PST')
    parser.add_argument('input', help='Archivo OST de entrada')
    parser.add_argument('output', help='Archivo PST de salida')
    parser.add_argument('--method', choices=['readpst', 'libpff', 'aspose'],
                       default='readpst', help='Método de conversión')

    args = parser.parse_args()

    # Mapear a enum
    method_map = {
        'readpst': ConversionMethod.READPST,
        'libpff': ConversionMethod.LIBPFF,
        'aspose': ConversionMethod.ASPOSE
    }

    converter = OSTtoPSTConverter()

    def print_log(message, level="INFO"):
        print(f"[{level}] {message}")

    success, message = converter.convert(
        args.input,
        args.output,
        method_map[args.method],
        progress_callback=print_log
    )

    if success:
        print(f"\n✓ Conversión exitosa: {args.output}")
        return 0
    else:
        print(f"\n✗ Error: {message}")
        return 1


if __name__ == "__main__":
    import sys
    sys.exit(main())
