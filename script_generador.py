from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert

doc = DocxTemplate("doc_plantilla.docx")

df=pd.read_excel("data.xlsx")

for index, fila in df.iterrows():
    context={"Postulacion": fila["Postulacion"]}

    doc.render(context)

    # Guardar el documento como archivo .docx
    nombre_archivo= f"{fila['Postulacion']}.Ingeniería Electromecánica. Vázquez Mori, Edwin.docx"
    ruta_archivo= f"{"Generated_docs/Español"}/{nombre_archivo}"
    doc.save(ruta_archivo)

    # Convertir el archivo .docx a .pdf
    nombre_archivo_pdf = f"{fila['Postulacion']}.Ingeniería Electromecánica. Vázquez Mori, Edwin.pdf"
    ruta_archivo_pdf = f"Generated_docs/Español/{nombre_archivo_pdf}"
    convert(ruta_archivo, ruta_archivo_pdf)