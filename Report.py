from docxtpl import DocxTemplate
import pandas as pd
doc = DocxTemplate("Template.docx")
context = {
    'name' : "Safee",
    'age':  "43"
}
doc.render(context)
doc.save("generated_doc.docx")