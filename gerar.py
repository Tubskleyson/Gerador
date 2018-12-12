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

print('> Modelo do Certificado')
file_path = os.path.abspath(input('\n< '))

print('> Lista de valores')
list_path = os.path.abspath(input('\n< '))

print('> Pasta de sáida')
out_path = os.path.abspath(input('\n< '))

if '\\' != out_path[-1]: out_path+='\\'

nomes = sorted([i.split(';')[0].strip('\n') for i in open(list_path).readlines()[1:]])

modelo = Document(file_path)

run = get_run(modelo)

word = client.Dispatch('Word.Application')


j = 1

for i in nomes:

    print('> Gerano arquivo %d de %d' %(j,len(nomes)),end='\r')
    j+=1
    run.text = i

    modelo.save(out_path+i+'.docx')

    doc = word.Documents.Open(os.path.abspath(out_path+i + '.docx'))
    doc.SaveAs(out_path+ i + '.pdf', FileFormat=wdFormatPDF)
    doc.Close()

    os.remove(out_path+i+'.docx')


word.Quit()
