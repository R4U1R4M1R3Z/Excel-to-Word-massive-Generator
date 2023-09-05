# Generador de Documentos Automático

Este script en Python te permite generar automáticamente documentos Word personalizados a partir de una plantilla y datos de un archivo Excel. Puede ser útil para crear documentos en masa con información específica para cada registro en el archivo Excel.

## Descripción

Este script automatiza el proceso de generación de documentos Word a partir de una plantilla predefinida y datos proporcionados en un archivo Excel. Puedes utilizar este script para ahorrar tiempo en la creación manual de documentos repetitivos, como certificados, informes o cualquier otro tipo de documento donde solo cambian ciertos valores.

## Requisitos

- Python 3.x instalado
- Bibliotecas: `pandas`, `docxtpl`

Puedes instalar las bibliotecas requeridas usando el siguiente comando:

```bash
pip install -r requirements.txt
```
## Cómo Funciona el Script

1. Cargar el archivo Excel: El script carga un archivo Excel que contiene los datos a utilizar para rellenar la plantilla de Word. Debes asegurarte de que el archivo Excel tenga las columnas necesarias (FURGON, CHASIS, EQUIPO_FRIO, NUM_SERIE) con los datos correspondientes.

2. Definir las Rutas de la Plantilla(**template_path**) y la Carpeta de Salida(**output_folder**): Debes especificar la ubicación de la plantilla de Word y la carpeta donde se guardarán los documentos generados. Asegúrate de que estas rutas sean correctas y accesibles.

3. Iterar a Través de los Registros en el DataFrame: El script itera a través de cada fila en el archivo Excel y obtiene los valores de las columnas en mi ejemplo(FURGON, CHASIS, EQUIPO_FRIO, NUM_SERIE) para cada registro.

4. Cargar la Plantilla de Word: Utiliza la biblioteca docxtpl para cargar la plantilla de Word como una plantilla que se puede llenar con datos.

5. Definir el Diccionario de Contexto: Crea un diccionario llamado **context** donde las claves son los nombres de las variables en la plantilla de Word y los valores son los datos que se insertarán en lugar de esas variables.

6. Renderizar la Plantilla: Utiliza doc.render(context) para reemplazar las variables en la plantilla de Word con los datos del diccionario de contexto.

7. Guardar el Documento Generado: El documento Word generado se guarda con un nombre de archivo único en la carpeta de salida.

10. Abrir la Carpeta con los Archivos Generados: Finalmente, el script utiliza os.system para abrir la carpeta que contiene los documentos generados.

### Cómo Poner las Variables en la Plantilla

Para insertar variables en la plantilla de Word que el script pueda reemplazar, debes seguir este formato:

En la plantilla de Word, encierra el nombre de la variable entre llaves dobles {{}}. Por ejemplo: {{FURGON}}, {{CHASIS}}, {{EQUIPO_FRIO}}, {{NUM_SERIE}}.

## Configuración

Edita el archivo `automa.py` para configurar las rutas y nombres de archivo:

- `excel_file`: Ruta del archivo Excel con los datos.
- `template_path`: Ruta de la plantilla de Word (.docx).
- `output_folder`: Carpeta de salida para los documentos generados.

## Contribuciones

Si encuentras algún error o tienes sugerencias de mejora, ¡no dudes en abrir un problema o enviar una solicitud de extracción!
