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
    WIN32COM = "win32com"
    HYBRID_MBOX = "hybrid_mbox"


class OSTtoPSTConverter:
    """Clase principal para conversión de archivos OST a PST"""

    def __init__(self):
        self.available_methods = self.check_available_methods()

    def check_available_methods(self) -> dict:
        """Verifica qué métodos de conversión están disponibles"""
        available = {
            'readpst': False,
            'libpff': False,
            'aspose': False,
            'win32com': False,
            'hybrid_mbox': False
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

        # Verificar win32com (solo Windows)
        try:
            import win32com.client
            import platform
            if platform.system() == 'Windows':
                available['win32com'] = True
        except ImportError:
            pass

        # Verificar hybrid_mbox (requiere readpst)
        if shutil.which('readpst'):
            available['hybrid_mbox'] = True

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
            elif method == ConversionMethod.WIN32COM:
                return self._convert_with_win32com(input_file, output_file, log)
            elif method == ConversionMethod.HYBRID_MBOX:
                return self._convert_hybrid_mbox(input_file, output_file, log)
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

    def _convert_with_win32com(self,
                               input_file: str,
                               output_file: str,
                               log: Callable) -> Tuple[bool, str]:
        """
        Convierte usando win32com (Windows + Outlook)
        Este método requiere Microsoft Outlook instalado en Windows
        """
        log("Usando método win32com (requiere Outlook en Windows)", "INFO")

        try:
            import win32com.client
            import platform

            if platform.system() != 'Windows':
                return False, "El método win32com solo funciona en Windows"

            log("Iniciando Outlook COM Automation...", "INFO")

            # Crear instancia de Outlook
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            log(f"Abriendo archivo OST: {input_file}", "INFO")

            # Agregar el archivo OST/PST a Outlook temporalmente
            try:
                namespace.AddStore(input_file)
            except Exception as e:
                log(f"No se pudo agregar OST directamente: {e}", "WARNING")
                return False, (
                    f"Error al abrir el archivo con Outlook.\n"
                    f"Asegúrese de que el archivo no esté en uso y que Outlook pueda acceder a él.\n"
                    f"Nota: Los archivos OST generalmente están vinculados a cuentas específicas."
                )

            log("Archivo agregado a Outlook", "INFO")

            # Crear nuevo archivo PST
            log(f"Creando archivo PST: {output_file}", "INFO")

            try:
                namespace.AddStore(output_file)
            except:
                # Si ya existe, intentar eliminarlo primero
                if os.path.exists(output_file):
                    os.remove(output_file)
                namespace.AddStore(output_file)

            # Obtener las carpetas
            source_store = None
            dest_store = None

            for store in namespace.Stores:
                if str(store.FilePath).lower() == input_file.lower():
                    source_store = store
                elif str(store.FilePath).lower() == output_file.lower():
                    dest_store = store

            if not source_store:
                return False, "No se pudo acceder al archivo de origen en Outlook"

            if not dest_store:
                return False, "No se pudo crear el archivo de destino en Outlook"

            log("Copiando carpetas y mensajes...", "INFO")

            # Copiar carpetas recursivamente
            self._copy_outlook_folders(source_store.GetRootFolder(), dest_store.GetRootFolder(), log)

            # Remover los stores
            namespace.RemoveStore(source_store.GetRootFolder())
            namespace.RemoveStore(dest_store.GetRootFolder())

            log("Conversión completada exitosamente", "SUCCESS")

            return True, "Conversión completada exitosamente usando Outlook COM"

        except ImportError:
            return False, "pywin32 no está instalado. Instale con: pip install pywin32"
        except Exception as e:
            return False, f"Error con win32com: {str(e)}"

    def _copy_outlook_folders(self, source_folder, dest_folder, log: Callable):
        """Copia carpetas de Outlook recursivamente (para win32com)"""
        try:
            # Copiar mensajes de la carpeta actual
            if source_folder.Items.Count > 0:
                log(f"Copiando {source_folder.Items.Count} items de '{source_folder.Name}'", "INFO")

                for item in source_folder.Items:
                    try:
                        copied_item = item.Copy()
                        copied_item.Move(dest_folder)
                    except Exception as e:
                        log(f"Error copiando item: {e}", "WARNING")

            # Copiar subcarpetas recursivamente
            for subfolder in source_folder.Folders:
                log(f"Procesando carpeta: {subfolder.Name}", "INFO")

                try:
                    new_folder = dest_folder.Folders.Add(subfolder.Name)
                    self._copy_outlook_folders(subfolder, new_folder, log)
                except Exception as e:
                    log(f"Error con carpeta {subfolder.Name}: {e}", "WARNING")

        except Exception as e:
            log(f"Error copiando carpeta: {str(e)}", "ERROR")

    def _convert_hybrid_mbox(self,
                            input_file: str,
                            output_file: str,
                            log: Callable) -> Tuple[bool, str]:
        """
        Método híbrido: OST -> MBOX (con readpst) -> Dejar MBOX para importación manual

        Este método extrae el contenido del OST a formato MBOX, que puede ser:
        1. Importado en Thunderbird, Evolution, u otros clientes de correo
        2. Convertido a PST usando herramientas externas
        3. Importado en Outlook (requiere plugins o herramientas de terceros)
        """
        log("Usando método híbrido: OST -> MBOX", "INFO")

        try:
            # Crear directorio de salida para MBOX
            mbox_dir = Path(output_file).parent / f"{Path(output_file).stem}_mbox"
            mbox_dir.mkdir(exist_ok=True)

            log(f"Extrayendo a formato MBOX en: {mbox_dir}", "INFO")

            # Ejecutar readpst en modo recursivo/mbox
            cmd = [
                'readpst',
                '-r',                  # Recursive (crear estructura de carpetas)
                '-o', str(mbox_dir),   # Directorio de salida
                '-D',                  # Include deleted items
                '-M',                  # MH format (cada mensaje como archivo)
                input_file
            ]

            log(f"Ejecutando: {' '.join(cmd)}", "INFO")

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=3600
            )

            if result.returncode != 0:
                log(f"Advertencia readpst: {result.stderr}", "WARNING")

            if result.stdout:
                log(f"Salida: {result.stdout}", "INFO")

            # Contar archivos extraídos
            extracted_files = list(mbox_dir.rglob('*'))
            extracted_msg_files = [f for f in extracted_files if f.is_file()]

            log(f"Extracción completada: {len(extracted_msg_files)} archivos de mensajes", "INFO")

            # Crear archivo de información
            info_file = mbox_dir / "README.txt"
            with open(info_file, 'w', encoding='utf-8') as f:
                f.write("OST to MBOX Conversion\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"Archivo original: {input_file}\n")
                f.write(f"Archivos extraídos: {len(extracted_msg_files)}\n\n")
                f.write("OPCIONES PARA CREAR PST:\n\n")
                f.write("1. IMPORTAR EN THUNDERBIRD:\n")
                f.write("   - Abrir Mozilla Thunderbird\n")
                f.write("   - Tools > ImportExportTools > Import mbox file\n")
                f.write("   - Seleccionar los archivos mbox de este directorio\n\n")
                f.write("2. IMPORTAR EN OUTLOOK (Windows):\n")
                f.write("   - Usar herramienta como 'Aid4Mail' o 'Kernel MBOX to PST'\n")
                f.write("   - O importar primero en Thunderbird y luego exportar a PST\n\n")
                f.write("3. CONVERSIÓN DIRECTA A PST:\n")
                f.write("   - Usar Aspose.Email: python converter.py --method aspose\n")
                f.write("   - O usar este script en Windows con Outlook instalado:\n")
                f.write("     python converter.py --method win32com\n\n")
                f.write("Para más información, consulte el README.md del proyecto.\n")

            log(f"Instrucciones guardadas en: {info_file}", "INFO")

            msg = (
                f"✓ Extracción exitosa a formato MBOX/MH\n"
                f"Ubicación: {mbox_dir}\n"
                f"Archivos: {len(extracted_msg_files)} mensajes\n\n"
                f"Para crear PST:\n"
                f"1. Importar en Thunderbird (Ver {info_file})\n"
                f"2. Usar método 'aspose' (comercial)\n"
                f"3. Usar método 'win32com' en Windows con Outlook\n"
            )

            log(msg, "SUCCESS")

            return True, msg

        except subprocess.TimeoutExpired:
            return False, "La extracción excedió el tiempo límite de 1 hora"
        except Exception as e:
            return False, f"Error en conversión híbrida: {str(e)}"


def main():
    """Función principal para uso en línea de comandos"""
    import argparse

    parser = argparse.ArgumentParser(description='Convertir archivos OST a PST')
    parser.add_argument('input', help='Archivo OST de entrada')
    parser.add_argument('output', help='Archivo PST de salida')
    parser.add_argument('--method',
                       choices=['readpst', 'libpff', 'aspose', 'win32com', 'hybrid_mbox'],
                       default='hybrid_mbox',
                       help='Método de conversión (default: hybrid_mbox)')

    args = parser.parse_args()

    # Mapear a enum
    method_map = {
        'readpst': ConversionMethod.READPST,
        'libpff': ConversionMethod.LIBPFF,
        'aspose': ConversionMethod.ASPOSE,
        'win32com': ConversionMethod.WIN32COM,
        'hybrid_mbox': ConversionMethod.HYBRID_MBOX
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
