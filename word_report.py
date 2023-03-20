import os, sys #Standard Python Libraries
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu

#Change path to current working directory
os.chdir(sys.path[0])

doc = DocxTemplate('Template_2.docx')
placeholder_1 = InlineImage(doc, 'Placeholders/Placeholder.png', Cm(5))
context = {
    'name' : 'Pablo',
    'placeholder_1' : placeholder_1,
    'placeholder_2' : placeholder_1}

doc.render(context)
doc.save('Template_Rendered_2.docx')