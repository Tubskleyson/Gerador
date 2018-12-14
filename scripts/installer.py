from zipfile import ZipFile
import os
from tkinter import *
from PIL import ImageTk, Image
import winshell

class Installer:

    def __init__(self):

        self.window = Tk()
        self.window.iconbitmap('assets/install.ico')

        self.window.title('Installer')

        self.window.resizable(False, False)

        self.window.geometry('550x300')

        self.window.configure(background='white')

        load = Image.open("assets/ins.png")

        space = Label(self.window, bg='white')
        space.pack(side=BOTTOM, pady=2)

        load = load.resize((180, 180), Image.ANTIALIAS)

        img = ImageTk.PhotoImage(load)

        image = Label(self.window, text='kalfoi', image=img, bg = 'white')
        image.image = img
        image.pack(side=LEFT,padx=20)

        space = Label(self.window, bg='white')
        space.pack(pady=20)



        tit = Label(self.window, text='Instalador                   ', bg='white', font="sans 25 ", fg='grey')
        tit.pack()

        self.op = Label(self.window, text='Selecione as opções desejadas\t     \n', bg='white', font="sans 13 ", fg='grey')
        self.op.pack()

        self.v = IntVar()
        c = Checkbutton(self.window, text="Criar ícone na área de trabalho\t\t\t\t", bg = 'white',onvalue=1,offvalue=0,variable=self.v)
        c.var = self.v
        c.pack()

        self.w = IntVar()
        d = Checkbutton(self.window, text="Adicionar ao menu Iniciar \t\t\t\t\t", bg='white', onvalue=1,offvalue=0,variable=self.w)
        d.var = self.w
        d.pack()

        self.bt = Button(self.window, text='Instalar', height=2, width=10, cursor='hand2', command=self.load)
        self.bt.pack(side=RIGHT,padx=30)

        self.widgets = [c,d]

    def load(self):
        self.bt['cursor'] = 'watch'
        self.window['cursor'] = 'watch'
        self.window.after(10,self.start)

    def start(self):

        user_path = '\\'.join([ i for i in os.getcwd().split('\\')[:3]]) + '\\'


        self.zip = ZipFile('Tubs.zip')

        self.zip.extractall(user_path)

        self.path = user_path + 'Tubs\\'

        lista = [(self.v.get(), winshell.desktop()),(self.w.get(),os.path.abspath(user_path+"\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs"))]

        for i in lista:

            if i[0]:

                link_filepath = os.path.join(i[1], "Gerador.lnk")
                with winshell.shortcut(link_filepath) as link:
                    link.path = self.path+'interface.exe'
                    link.description = "Gerador de Certificados"
                    link.working_directory = self.path



        self.window['cursor'] = 'arrow'
        self.bt['cursor'] = 'hand2'

        for i in self.widgets: i.destroy()

        self.op['text'] = 'Instalação Finalizada\t\t    '

        self.bt['text'] = 'Encerrar'
        self.bt['command'] = self.window.destroy




app = Installer()
app.window.mainloop()