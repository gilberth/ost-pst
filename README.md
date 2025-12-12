# OST to PST Converter

Una aplicación con interfaz gráfica para convertir archivos OST (Offline Storage Table) a PST (Personal Storage Table) de Microsoft Outlook.

## Características

- **Interfaz gráfica amigable** construida con tkinter
- **5 métodos de conversión diferentes**:
  - **Híbrido MBOX** (código abierto, RECOMENDADO) - Extrae a MBOX para importación
  - **Aspose.Email** (comercial) - Conversión directa OST→PST
  - **Win32COM** (Windows + Outlook) - Usa Outlook para conversión nativa
  - **readpst** (código abierto) - Extracción de contenido
  - **libpff** (código abierto) - Solo lectura de datos
- **Registro detallado** del proceso de conversión
- **Validación de archivos** de entrada y salida
- **Multiplataforma** (Linux, Windows, macOS)
- **Sin limitaciones artificiales** en archivos gratuitos

## Capturas de pantalla

```
┌─────────────────────────────────────────────┐
│         Convertidor OST a PST               │
├─────────────────────────────────────────────┤
│                                             │
│ Archivo OST de Entrada                      │
│ ┌──────────────────────┐ [Examinar...]     │
│ │ /ruta/al/archivo.ost │                    │
│ └──────────────────────┘                    │
│                                             │
│ Archivo PST de Salida                       │
│ ┌──────────────────────┐ [Examinar...]     │
│ │ /ruta/al/salida.pst  │                    │
│ └──────────────────────┘                    │
│                                             │
│ Opciones de Conversión                      │
│ ○ readpst  ○ libpff  ○ Aspose.Email       │
│                                             │
│       [Convertir] [Cancelar] [Salir]       │
│                                             │
│ Registro de Conversión                      │
│ ┌─────────────────────────────────────────┐ │
│ │ [INFO] Iniciando conversión...          │ │
│ │ [INFO] Método: readpst                  │ │
│ │ ...                                     │ │
│ └─────────────────────────────────────────┘ │
│                                             │
│ [████████████████] 100%                     │
│ Estado: Listo                               │
└─────────────────────────────────────────────┘
```

## Requisitos del sistema

### Python
- Python 3.5 o superior
- tkinter (generalmente incluido con Python)

### Herramientas de conversión (al menos una):

#### Opción 1: readpst (Recomendado - Código abierto)
```bash
# Debian/Ubuntu
sudo apt-get install pst-utils

# RedHat/Fedora
sudo dnf install pst-utils

# macOS
brew install readpst
```

#### Opción 2: libpff con pypff (Código abierto)
```bash
# Debian/Ubuntu
sudo apt-get install libpff-dev libpff-python3
pip install pypff

# macOS
brew install libpff
pip install pypff
```

#### Opción 3: Aspose.Email (Comercial - Requiere licencia)
```bash
pip install Aspose.Email-for-Python-via-NET
```
**Nota**: Aspose.Email es una biblioteca comercial que requiere una licencia para uso en producción.

## Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/gilberth/ost-pst.git
cd ost-pst
```

### 2. Instalar dependencias del sistema
```bash
# Opción recomendada: readpst
sudo apt-get install pst-utils

# O usar el script de instalación automático
chmod +x install.sh
./install.sh
```

### 3. Instalar dependencias de Python (opcional)
```bash
pip install -r requirements.txt
```

## Uso

### Interfaz gráfica

```bash
python3 app.py
```

Luego:
1. Haz clic en **"Examinar..."** junto a "Archivo OST de Entrada" y selecciona tu archivo .ost
2. Haz clic en **"Examinar..."** junto a "Archivo PST de Salida" y elige dónde guardar el archivo .pst
3. Selecciona el **método de conversión** preferido
4. Haz clic en **"Convertir"**
5. Espera a que se complete el proceso (puedes ver el progreso en el registro)

### Línea de comandos

```bash
python3 converter.py input.ost output.pst --method readpst
```

Opciones:
- `input.ost`: Archivo OST de entrada
- `output.pst`: Archivo PST de salida
- `--method`: Método de conversión (`hybrid_mbox`, `aspose`, `win32com`, `readpst`, `libpff`)

### Ejemplos

```bash
# Conversión con método híbrido MBOX (RECOMENDADO - Gratuito)
python3 converter.py /path/to/mailbox.ost /path/to/output.pst --method hybrid_mbox

# Usar Aspose.Email (requiere licencia comercial)
python3 converter.py mailbox.ost output.pst --method aspose

# Usar Win32COM en Windows (requiere Outlook instalado)
python converter.py mailbox.ost output.pst --method win32com

# Solo extracción con readpst
python3 converter.py mailbox.ost output.pst --method readpst
```

## Métodos de conversión

### 🌟 Híbrido MBOX (RECOMENDADO - Sin limitaciones)
- ✅ **100% Código abierto y GRATUITO**
- ✅ **Sin limitaciones** - convierte archivos de cualquier tamaño
- ✅ Extrae todo el contenido a formato MBOX/MH
- ✅ Compatible con Thunderbird, Evolution y otros clientes
- ✅ Los archivos MBOX pueden importarse a PST posteriormente
- ℹ️ **Mejor opción gratuita sin restricciones**
- 📝 Incluye instrucciones detalladas para el siguiente paso

**Flujo**: OST → MBOX → (importar en Thunderbird/Outlook)

### Aspose.Email
- ✅ **Conversión directa** OST → PST
- ✅ Muy robusto y completo
- ✅ Soporta versiones modernas de OST/PST
- ✅ No requiere Outlook instalado
- ❌ **Comercial** - requiere licencia (~$1000 USD)
- ⚠️ No soporta formatos OST 2013/2016 en algunas versiones
- ℹ️ Mejor para entornos empresariales con presupuesto

### Win32COM (Windows + Outlook)
- ✅ **Conversión directa** OST → PST
- ✅ **Gratuito** (usa Outlook instalado)
- ✅ Conversión nativa de Microsoft
- ❌ **Solo Windows** con Outlook instalado
- ⚠️ Los archivos OST pueden estar vinculados a cuentas específicas
- 📦 Requiere: `pip install pywin32`
- ℹ️ Mejor para usuarios de Windows que ya tienen Outlook

### readpst (Solo extracción)
- ✅ **Código abierto** y gratuito
- ✅ Ampliamente probado y mantenido
- ✅ Extrae a múltiples formatos (MBOX, MH, EML)
- ⚠️ **No crea PST nativos** - solo extrae contenido
- ℹ️ Usado internamente por el método híbrido MBOX

### libpff (Solo lectura)
- ✅ **Código abierto** y gratuito
- ✅ Bueno para análisis forense
- ❌ **Solo lectura** - no puede escribir PST
- ℹ️ Mejor para inspección y extracción de datos específicos

## Comparación de métodos

| Método | Costo | Limitaciones | Plataforma | PST Nativo | Recomendación |
|--------|-------|--------------|------------|------------|---------------|
| **Híbrido MBOX** | Gratis | Ninguna | Todas | Vía importación | ⭐⭐⭐⭐⭐ **MEJOR OPCIÓN GRATUITA** |
| **Aspose.Email** | ~$1000 | Requiere licencia | Todas | Sí | ⭐⭐⭐⭐ Empresas |
| **Win32COM** | Gratis | Solo Windows+Outlook | Windows | Sí | ⭐⭐⭐⭐ Si tienes Outlook |
| **readpst** | Gratis | No crea PST | Todas | No | ⭐⭐⭐ Solo extracción |
| **libpff** | Gratis | Solo lectura | Todas | No | ⭐⭐ Solo análisis |

## Notas importantes

1. **Método Híbrido MBOX** es la mejor opción sin costo:
   - Extrae TODO el contenido sin limitaciones
   - Los archivos MBOX pueden importarse en Thunderbird, Evolution, etc.
   - Desde Thunderbird puedes exportar a PST con plugins
   - Sin restricciones de tamaño o cantidad de mensajes

2. **Archivos OST vinculados**: Algunos archivos OST están vinculados a cuentas de Exchange específicas y pueden requerir autenticación.

3. **Archivos muy grandes**: La conversión de archivos OST de varios GB puede tomar tiempo considerable (normal).

## Solución de problemas

### "No se encontraron herramientas de conversión instaladas"
Instala al menos readpst para usar el método híbrido MBOX (recomendado):
```bash
sudo apt-get install pst-utils  # Debian/Ubuntu
```

### "readpst: command not found"
```bash
# Debian/Ubuntu
sudo apt-get install pst-utils

# macOS
brew install readpst
```

### "win32com no disponible"
Solo en Windows:
```bash
pip install pywin32
```
Asegúrate de tener Microsoft Outlook instalado.

### "Error: pypff no está instalado"
```bash
sudo apt-get install libpff-dev libpff-python3
pip install pypff
```

### "Error de permisos"
Asegúrate de tener permisos de lectura en el archivo OST y de escritura en el directorio de salida.

### "El archivo OST está vinculado a una cuenta"
Los archivos OST de Exchange suelen estar vinculados a cuentas específicas. Opciones:
1. Usa el método **híbrido MBOX** que extrae el contenido sin necesidad de la cuenta
2. Si tienes acceso a la cuenta, usa **Win32COM** en Windows con Outlook configurado

### Conversión muy lenta
Los archivos OST grandes pueden tardar mucho tiempo. Esto es normal. Verifica el registro para asegurarte de que el proceso está avanzando.

### ¿Cómo convertir MBOX a PST?
Después de usar el método híbrido MBOX:
1. **Thunderbird** (gratuito):
   - Importa los archivos MBOX en Thunderbird
   - Usa plugin ImportExportTools NG
   - Opcional: exporta a PST con otros plugins

2. **Outlook en Windows**:
   - Importa los MBOX con herramientas como Aid4Mail
   - O usa el método win32com directamente

3. **Conversión directa**:
   - Usa el método Aspose.Email (comercial)
   - Usa el método Win32COM si estás en Windows

## Estructura del proyecto

```
ost-pst/
├── app.py              # Aplicación con interfaz gráfica
├── converter.py        # Lógica de conversión (puede usarse standalone)
├── requirements.txt    # Dependencias de Python
├── install.sh         # Script de instalación para Linux
├── README.md          # Este archivo
└── .gitignore         # Archivos a ignorar en git
```

## Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Haz fork del repositorio
2. Crea una rama para tu funcionalidad (`git checkout -b feature/nueva-funcionalidad`)
3. Haz commit de tus cambios (`git commit -am 'Agrega nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Crea un Pull Request

## Licencia

Este proyecto es de código abierto y está disponible bajo la licencia MIT.

## Advertencias y consideraciones

- **Privacidad**: Los archivos OST/PST contienen correos electrónicos y datos personales. Asegúrate de manejarlos de manera segura.
- **Backups**: Siempre mantén un respaldo del archivo OST original antes de intentar la conversión.
- **Validación**: Después de la conversión, verifica que el archivo PST resultante se abre correctamente en Outlook.
- **Licencias**: Si usas Aspose.Email en producción, asegúrate de tener la licencia apropiada.

## Soporte

Para reportar problemas o solicitar funcionalidades, por favor abre un issue en el repositorio de GitHub.

## Referencias y documentación

- [readpst Documentation](http://www.five-ten-sg.com/libpst/)
- [libpff Project](https://github.com/libyal/libpff)
- [Aspose.Email Documentation](https://docs.aspose.com/email/python-net/)

## Autor

Desarrollado para facilitar la conversión de archivos OST a PST sin depender exclusivamente de Microsoft Outlook.

---

**Nota importante**: Esta herramienta es para uso legítimo con tus propios archivos OST o archivos para los cuales tienes autorización. Respeta siempre la privacidad y las leyes de protección de datos aplicables.
