#!/bin/bash

# Exclude files and folders listed in .exclude from build/dist
EXCLUDE_FILE=".exclude"

# Name of the output executable
OUTPUT="FacturaElectronica"

# Find main Python file (adjust if needed)
MAIN="ui.py"

# Instalar dependencias si existe requirements.txt
if [ -f "requirements.txt" ]; then
	pip install -r requirements.txt
	# Generar opciones --hidden-import para cada paquete
	HIDDEN_IMPORTS=""
	while read -r pkg; do
		# Ignorar líneas vacías y comentarios
		if [[ ! "$pkg" =~ ^# ]] && [[ -n "$pkg" ]]; then
			# Extraer solo el nombre del paquete (sin versión)
			pkg_name=$(echo "$pkg" | cut -d'=' -f1 | cut -d'>' -f1 | cut -d'<' -f1)
			HIDDEN_IMPORTS="$HIDDEN_IMPORTS --hidden-import=$pkg_name"
		fi
	done < requirements.txt
else
	HIDDEN_IMPORTS=""
fi

# Compilar con PyInstaller del entorno activo incluyendo hidden-imports y submódulos de PySide6
"$(which pyinstaller)" --onefile --name "$OUTPUT" $EXCLUDE $HIDDEN_IMPORTS \
	--hidden-import=PySide6 \
	--hidden-import=PySide6.QtCore \
	--hidden-import=PySide6.QtGui \
	--hidden-import=PySide6.QtWidgets \
	--hidden-import=PySide6.QtNetwork \
	--hidden-import=PySide6.QtPrintSupport \
	--hidden-import=rich.progress \
	--hidden-import=rich.console \
	--hidden-import=openpyxl.utils \
	--hidden-import=openpyxl.cell._writer \
	"$MAIN"
