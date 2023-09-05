import os
import pandas as pd
from docxtpl import DocxTemplate

# Cargar el archivo Excel
excel_file = "D:\\script\\data.xlsx"
df = pd.read_excel(excel_file)

# Ruta de la plantilla de Word y carpeta de salida para los documentos generados
template_path = "D:\\script\\template.docx"
output_folder = "D:\\script\\documentos"

# Iterar a través de los registros en el DataFrame
for index, row in df.iterrows():
    furgon = row["FURGON"]
    chasis = row["CHASIS"]
    equipo_frio = row["EQUIPO_FRIO"]
    NUM_SERIE = row["NUM_SERIE"]

    # Cargar la plantilla de Word como una plantilla docxtpl
    doc = DocxTemplate(template_path)

    # Definir los datos a reemplazar en la plantilla
    context = {
        'FURGON': furgon,
        'CHASIS': chasis,
        'EQUIPO_FRIO': equipo_frio,
        'NUM_SERIE': NUM_SERIE
    }

    # Renderizar la plantilla con los datos
    doc.render(context)
    
    # Guardar el documento Word generado
    output_filename = f"{output_folder}\\Certificado - {chasis}.docx"
    doc.save(output_filename)

print("Documentos generados correctamente.")

# Abrir la carpeta con los archivos generados
os.system(f'explorer "{output_folder}"')


print("Documentos generados correctamente.")

# Abrir la carpeta con los archivos generados
os.system(f'explorer "{output_folder}"')
