from PyPDF2 import PdfFileReader,PdfFileWriter
import os

class App:

    def __init__(self):

        print("\n> Bem vindo ao separador de pdf.\n")

        self.path = ''
        self.filename = ''
        self.pdf = None
        self.lista = ''

    def run(self):

        while not self.path: self.path = self.get_path()

        os.chdir(self.path)

        while not self.filename: self.filename = self.get_filename()

        self.pdf = PdfFileReader(open(self.filename, 'rb'))

        self.paginas = self.pdf.numPages

        print("\n> Você deseja basear o nome dos arquivos gerados numa lista?")
        r = input("< ")

        while r not in ['s', 'n', 'sim', 'não', 'nao']:
            print("\n> Desculpse, não entendi sua resposta, você pode dizer sim ou não")
            r = input("< ")

        if r=='sim' or r=='s': self.lista = self.get_lista()

        print("> Ok, digite enter para começar a separação")
        input()

        self.gerar()

        print("\n> Tudo certo. Você vai achar seus %d pdfs na pasta %s" % (self.paginas, self.out_path))

        input()


    def get_path(self):

        print("\n> Em que pasta está o pdf que você deseja dividir?")
        path = input("\n< ")

        if os.path.exists(path):
            if '\\' != path[-1]: path += '\\'
            return path

        else:

            print("\n> Caminho inválido\n")
            return 0


    def get_filename(self):

        print("\n> Qual o nome do arquivo a ser dividido?")
        filename = input("\n< ")

        if '.pdf' not in filename: filename += '.pdf'

        if os.path.exists(filename): return filename

        else:
            print("\n> Não encontrei o arquivo %s" %filename)
            return 0


    def get_lista(self):

        print("\n> E qual a pasta dessa lista?")
        lpath = input("\n< ")

        while not os.path.exists(lpath):
            print("\n> Não achei essa pasta, tente novamente")
            lpath = input("\n< ")

        if '\\' != lpath[-1]: lpath += '\\'

        print("\n> Ok, qual o nome do arquivo da lista? Note que eu só aceito arquivos .csv")

        listname = input("\n< ")

        if '.csv' not in listname: listname += '.csv'

        while not os.path.exists(listname):

            print("\n> Não consigo achar sua lista. Cheque o nome e tente novamente")
            listname = input("\n< ")

            if '.csv' not in listname: listname += '.csv'

        myl = open(lpath + listname)

        return sorted([i.strip('\n').split(';')[0] for i in myl.readlines()[1:]])

    def gerar(self):

        self.out_path = self.path + 'output\\'

        if not os.path.exists(self.out_path): os.mkdir(self.out_path)

        for i in range(self.paginas):

            output = PdfFileWriter()
            output.addPage(self.pdf.getPage(i))

            if self.lista:
                outFile = open(self.out_path + self.lista[i] + '.pdf', "wb")
            else:
                outFile = open(self.out_path + self.filename.split('.')[0] + str(i) + '.pdf', "wb")

            output.write(outFile)

