from PyPDF2 import PdfFileReader,PdfFileWriter
import os

print("\n> Bem vindo ao separador de pdf.\n")

path = 0

while not path:

    print("\n> Em que pasta está o pdf que você deseja dividir?")
    path = input("\n< ")

    try:
        os.chdir(path)
    except:

        print("\n> Caminho inválido\n")
        path = 0

if '/'!=path[-1]: path += '/'

print("\n> Ok, qual o nome do arquivo?")

filename = input("\n< ")

if '.pdf' not in filename: filename+='.pdf'

while not os.path.exists(filename):

    print("\n> Não consigo achar seu arquivo. Cheque o nome e tente novamente")
    filename = input("\n< ")

    if '.pdf' not in filename: filename += '.pdf'


pdf = PdfFileReader(open(filename,'rb'))
paginas = pdf.numPages


print("\n> Você deseja basear o nome dos arquivos gerados numa lista?")
r = input("< ")

while r not in ['s','n','sim','não','nao']:

    print("\n> Desculpse, não entendi sua resposta, você pode dizer sim ou não")
    r = input("< ")


if r=='sim':


    print("\n> E qual a pasta dessa lista?")
    lpath = input("\n< ")

    while not os.path.exists(lpath):

        print("\n> Não achei essa pasta, tente novamente")
        lpath = input("\n< ")

    if '/'!=lpath[-1]: lpath += '/'


    print("\n> Ok, qual o nome do arquivo da lista? Note que eu só aceito arquivos .csv")

    listname = input("\n< ")

    if '.csv' not in listname: listname += '.csv'

    while not os.path.exists(listname):

        print("\n> Não consigo achar sua lista. Cheque o nome e tente novamente")
        listname = input("\n< ")

        if '.csv' not in listname: listname += '.csv'

    myl = open(lpath+listname)

    lista = sorted([i.strip('\n').split(';')[0] for i in myl.readlines()[1:]])

else: lista = 0


print("> Ok, estou começando a separar as paginas")

np = path+'output/'

if not os.path.exists(np): os.mkdir(np)



for i in range(paginas):

    output = PdfFileWriter()
    output.addPage(pdf.getPage(i))

    if lista:  outFile = open(np+lista[i]+'.pdf',"wb")
    else: outFile = open(np+filename.split('.')[0]+str(i)+'.pdf',"wb")


    output.write(outFile)

print("\n> Tudo certo. Você vai achar seus %d pdfs na pasta %s" %(paginas,np))

input()