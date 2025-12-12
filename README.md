# OST to PST Converter

Una aplicación con interfaz gráfica para convertir archivos OST (Offline Storage Table) a PST (Personal Storage Table) de Microsoft Outlook.

## Características

- **Interfaz gráfica amigable** construida con tkinter
- **Múltiples métodos de conversión**:
  - **readpst** (código abierto, recomendado)
  - **libpff** (código abierto)
  - **Aspose.Email** (comercial, más robusto)
- **Registro detallado** del proceso de conversión
- **Validación de archivos** de entrada y salida
- **Multiplataforma** (Linux, Windows, macOS)

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
- `--method`: Método de conversión (`readpst`, `libpff`, o `aspose`)

### Ejemplos

```bash
# Conversión básica con readpst
python3 converter.py /path/to/mailbox.ost /path/to/output.pst

# Usar libpff
python3 converter.py mailbox.ost output.pst --method libpff

# Usar Aspose.Email (requiere licencia)
python3 converter.py mailbox.ost output.pst --method aspose
```

## Métodos de conversión

### readpst (Recomendado)
- ✅ **Código abierto** y gratuito
- ✅ Ampliamente probado y mantenido
- ✅ Soporta la mayoría de versiones de OST/PST
- ⚠️ Extrae contenido pero puede requerir pasos adicionales para crear PST nativo

### libpff (pypff)
- ✅ **Código abierto** y gratuito
- ✅ Buen rendimiento
- ⚠️ Solo lectura - no puede escribir PST directamente
- ℹ️ Mejor para análisis y extracción de datos

### Aspose.Email
- ✅ **Conversión nativa** OST → PST
- ✅ Muy robusto y completo
- ✅ Soporta todas las versiones de OST/PST
- ❌ **Comercial** - requiere licencia
- ℹ️ Mejor para entornos de producción empresariales

## Limitaciones conocidas

1. **readpst**: Extrae el contenido del OST pero puede no crear un archivo PST nativo directamente. Los archivos extraídos pueden necesitar importarse manualmente en Outlook.

2. **libpff/pypff**: Solo soporta lectura de archivos OST/PST, no escritura. No puede crear archivos PST nativos.

3. **Aspose.Email**:
   - Requiere licencia comercial
   - No soporta formatos OST 2013/2016 en algunas versiones (verificar documentación)

4. **Archivos muy grandes**: La conversión de archivos OST de varios GB puede tomar tiempo considerable y requerir memoria suficiente.

## Solución de problemas

### "No se encontraron herramientas de conversión instaladas"
Instala al menos una de las herramientas mencionadas en la sección de requisitos.

### "readpst: command not found"
```bash
sudo apt-get install pst-utils
```

### "Error: pypff no está instalado"
```bash
sudo apt-get install libpff-dev libpff-python3
pip install pypff
```

### "Error de permisos"
Asegúrate de tener permisos de lectura en el archivo OST y de escritura en el directorio de salida.

### Conversión muy lenta
Los archivos OST grandes pueden tardar mucho tiempo. Esto es normal. Verifica el registro para asegurarte de que el proceso está avanzando.

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
