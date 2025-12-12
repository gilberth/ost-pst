#!/bin/bash
# Script de instalación para OST to PST Converter
# Soporta Debian/Ubuntu, RedHat/Fedora, y macOS

set -e  # Salir si hay errores

echo "========================================="
echo "  OST to PST Converter - Instalación"
echo "========================================="
echo ""

# Detectar sistema operativo
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    # Detectar distribución Linux
    if [ -f /etc/debian_version ]; then
        DISTRO="debian"
        echo "✓ Sistema detectado: Debian/Ubuntu"
    elif [ -f /etc/redhat-release ]; then
        DISTRO="redhat"
        echo "✓ Sistema detectado: RedHat/Fedora/CentOS"
    else
        DISTRO="unknown"
        echo "⚠ Sistema Linux desconocido"
    fi
elif [[ "$OSTYPE" == "darwin"* ]]; then
    DISTRO="macos"
    echo "✓ Sistema detectado: macOS"
else
    DISTRO="unknown"
    echo "⚠ Sistema operativo no soportado: $OSTYPE"
fi

echo ""
echo "Instalando dependencias del sistema..."
echo ""

# Instalar dependencias según el sistema
case $DISTRO in
    debian)
        echo "Actualizando lista de paquetes..."
        sudo apt-get update

        echo ""
        echo "Instalando Python3 y tkinter..."
        sudo apt-get install -y python3 python3-tk python3-pip

        echo ""
        echo "Instalando pst-utils (readpst)..."
        sudo apt-get install -y pst-utils

        echo ""
        echo "¿Desea instalar libpff también? (s/n)"
        read -r install_libpff
        if [[ $install_libpff == "s" || $install_libpff == "S" ]]; then
            echo "Instalando libpff..."
            sudo apt-get install -y libpff-dev libpff-python3
        fi
        ;;

    redhat)
        echo "Instalando Python3 y tkinter..."
        sudo dnf install -y python3 python3-tkinter python3-pip

        echo ""
        echo "Instalando pst-utils (readpst)..."
        sudo dnf install -y pst-utils

        echo ""
        echo "¿Desea instalar libpff también? (s/n)"
        read -r install_libpff
        if [[ $install_libpff == "s" || $install_libpff == "S" ]]; then
            echo "Instalando libpff..."
            sudo dnf install -y libpff-devel
        fi
        ;;

    macos)
        # Verificar si Homebrew está instalado
        if ! command -v brew &> /dev/null; then
            echo "⚠ Homebrew no está instalado."
            echo "Por favor instala Homebrew desde https://brew.sh/"
            echo "Luego ejecuta este script nuevamente."
            exit 1
        fi

        echo "Instalando readpst con Homebrew..."
        brew install readpst

        echo ""
        echo "¿Desea instalar libpff también? (s/n)"
        read -r install_libpff
        if [[ $install_libpff == "s" || $install_libpff == "S" ]]; then
            echo "Instalando libpff..."
            brew install libpff
        fi
        ;;

    *)
        echo "⚠ Sistema no soportado para instalación automática."
        echo "Por favor instala manualmente:"
        echo "  - Python 3.5 o superior"
        echo "  - tkinter para Python"
        echo "  - pst-utils (readpst)"
        echo ""
        echo "¿Desea continuar con la instalación de dependencias Python? (s/n)"
        read -r continue_install
        if [[ $continue_install != "s" && $continue_install != "S" ]]; then
            exit 1
        fi
        ;;
esac

echo ""
echo "========================================="
echo "Instalando dependencias de Python (opcional)"
echo "========================================="
echo ""

# Instalar dependencias de Python
echo "¿Desea instalar pypff (Python libpff)? (s/n)"
read -r install_pypff
if [[ $install_pypff == "s" || $install_pypff == "S" ]]; then
    echo "Instalando pypff..."
    pip3 install pypff || echo "⚠ Error instalando pypff (puede requerir libpff instalado en el sistema)"
fi

echo ""
echo "¿Desea instalar Aspose.Email? (biblioteca comercial) (s/n)"
read -r install_aspose
if [[ $install_aspose == "s" || $install_aspose == "S" ]]; then
    echo "Instalando Aspose.Email..."
    echo "⚠ NOTA: Esta es una biblioteca comercial que requiere licencia."
    pip3 install Aspose.Email-for-Python-via-NET
fi

echo ""
echo "========================================="
echo "Configurando la aplicación"
echo "========================================="
echo ""

# Hacer ejecutables los scripts Python
chmod +x app.py
chmod +x converter.py

echo "✓ Scripts configurados como ejecutables"

echo ""
echo "========================================="
echo "Verificando instalación"
echo "========================================="
echo ""

# Verificar Python
if command -v python3 &> /dev/null; then
    PYTHON_VERSION=$(python3 --version)
    echo "✓ Python instalado: $PYTHON_VERSION"
else
    echo "✗ Python3 no encontrado"
fi

# Verificar readpst
if command -v readpst &> /dev/null; then
    echo "✓ readpst instalado"
else
    echo "✗ readpst no encontrado"
fi

# Verificar libpff
if command -v pffexport &> /dev/null; then
    echo "✓ libpff instalado"
else
    echo "✗ libpff no encontrado (opcional)"
fi

# Verificar pypff
if python3 -c "import pypff" 2>/dev/null; then
    echo "✓ pypff (Python) instalado"
else
    echo "✗ pypff no encontrado (opcional)"
fi

# Verificar Aspose
if python3 -c "import aspose.email" 2>/dev/null; then
    echo "✓ Aspose.Email instalado"
else
    echo "✗ Aspose.Email no encontrado (opcional)"
fi

echo ""
echo "========================================="
echo "¡Instalación completada!"
echo "========================================="
echo ""
echo "Para iniciar la aplicación, ejecuta:"
echo "  python3 app.py"
echo ""
echo "O para usar la línea de comandos:"
echo "  python3 converter.py input.ost output.pst"
echo ""
echo "Para más información, consulta el README.md"
echo ""
