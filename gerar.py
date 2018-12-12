import sys
import os
from win32com import client
from docx import Document

def get_run(modelo):

    for i in modelo.tables:
        x = len(i.columns)
        y = len(i.rows)

        for j in range(x):
            for k in range(y):
                ps = i.cell(k, j).paragraphs
                for l in ps:
                    if l.runs and l.runs[0].text == 'Valor_1':
                        return l.runs[0]

wdFormatPDF = 17

file_path = os.path.abspath(sys.argv[1])
list_path = os.path.abspath(sys.argv[2])

nomes = sorted([i.split(';')[0].strip('\n')[1:-1] for i in open(list_path).readlines()[1:]])

modelo = Document(file_path)

run = get_run(modelo)

word = client.Dispatch('Word.Application')


j = 1

for i in nomes:

    print('> Gerano arquivo %d de %d' %(j,len(nomes)),end='\r')
    j+=1
    run.text = i

    modelo.save(i+'.docx')

    doc = word.Documents.Open(os.path.abspath(i + '.docx'))
    doc.SaveAs(os.path.abspath('output\\'+ i + '.pdf'), FileFormat=wdFormatPDF)
    doc.Close()

    os.remove(i+'.docx')


word.Quit()