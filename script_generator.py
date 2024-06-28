from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert

doc = DocxTemplate("doc_template.docx")

df = pd.read_excel("data.xlsx")

for index, row in df.iterrows():
    context = {"Postulation": row["Postulation"]}

    doc.render(context)

    nameFile = f"{row['Postulation']}.Electromechanical engineering. Vázquez Mori, Edwin.docx"
    rootFile = f"Generated_docs/English/{nameFile}"
    doc.save(rootFile)

    nameFilePdf = f"{row['Postulation']}.Electromechanical engineering. Vázquez Mori, Edwin.pdf"
    rootFilePdf = f"Generated_docs/English/{nameFilePdf}"
    convert(rootFile, rootFilePdf)