# FacturaElectronica

Este programa permite procesar facturas electrónicas de manera automatizada. Está diseñado para extraer, organizar y manejar archivos relacionados con facturación electrónica.

## Funcionamiento

1. **Procesamiento de archivos:** El programa toma archivos de facturas electrónicas y los procesa, clasificando los documentos en carpetas como `extraidos` y `procesados`.
2. **Automatización:** El flujo de trabajo está automatizado para facilitar la gestión de grandes volúmenes de facturas.

## Ejecución del binario compilado

1. **Compilación:** Si aún no tienes el binario, compílalo usando el script correspondiente (`build.sh` en Linux o `build.bat` en Windows).
   ```bash
   ./build.sh
   ```
   o en Windows:
   ```bat
   build.bat
   ```

2. **Ejecución:** Una vez compilado, encontrarás el binario en la carpeta `dist`. Para ejecutarlo:
   ```bash
   ./dist/FacturaElectronica
   ```
   o en Windows:
   ```bat
   dist\FacturaElectronica.exe
   ```

3. **Parámetros:** Consulta la documentación interna del programa o ejecuta el binario con la opción `--help` para ver los parámetros disponibles.

## Requisitos

- Python (si usas el script de compilación)
- Dependencias especificadas en el proyecto

## Notas

- Asegúrate de tener los permisos necesarios para ejecutar scripts y binarios.

