# Generador de Documentos Automático

Este script en Python te permite generar automáticamente documentos Word personalizados a partir de una plantilla y datos de un archivo Excel. Puede ser útil para crear documentos en masa con información específica para cada registro en el archivo Excel.

## Descripción

Este script automatiza el proceso de generación de documentos Word a partir de una plantilla predefinida y datos proporcionados en un archivo Excel. Puedes utilizar este script para ahorrar tiempo en la creación manual de documentos repetitivos, como certificados, informes o cualquier otro tipo de documento donde solo cambian ciertos valores.

## Requisitos

- Python 3.x instalado
- Bibliotecas: `pandas`, `python-docx`

Puedes instalar las bibliotecas requeridas usando el siguiente comando:

```bash
pip install -r requirements.txt
```

## Uso

1. Coloca tus datos en un archivo Excel en el formato deseado.
2. Crea una plantilla de Word (.docx) con marcadores que deseas reemplazar, por ejemplo, `variable`.
3. Modifica el archivo `config.py` con las rutas y nombres de archivo correspondientes.
4. Ejecuta el script:

```bash
python automa.py
```

Los documentos generados se guardarán en la carpeta especificada.

## Configuración

Edita el archivo `config.py` para configurar las rutas y nombres de archivo:

- `excel_file`: Ruta del archivo Excel con los datos.
- `template_path`: Ruta de la plantilla de Word (.docx).
- `output_folder`: Carpeta de salida para los documentos generados.

## Contribuciones

Si encuentras algún error o tienes sugerencias de mejora, ¡no dudes en abrir un problema o enviar una solicitud de extracción!
