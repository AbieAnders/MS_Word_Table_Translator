import os
import docx
from translatepy import Translator

os.chdir('C:/Users/Ab/VsCode Projects/Projects/Table Translator PSG')
doc = docx.Document()
doc.save('Table Document.docx')
translator_object = Translator()

title = doc.add_heading('The Table')
table = doc.add_table(rows = 1, cols = 3)
header_row = table.rows[0].cells
header_row[0].text = 'S.No'
header_row[1].text = 'Water'
header_row[2].text = 'Fire'

data = (
    (1, 'yes', 'no'),
    (2, 'no', 'yes'),
)

for sno, choice1, choice2 in data:
    new_row = table.add_row()
    new_cells = new_row.cells
    new_cells[0].text = str(sno)
    new_cells[1].text = choice1
    new_cells[2].text = choice2
table.style = 'Colorful List'
for n,row in enumerate(table.rows):
    for m,cell in enumerate(row.cells):
        translation = translator_object.translate(cell.text,'ta')
        table.cell(n,m).text = str(translation)  #str type casting is necessary since the text function automatically assumes the TranslationResult type
doc.save('Table Document.docx')
