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

        op = Label(self.window, text='Selecione as opções desejadas\t     \n', bg='white', font="sans 13 ", fg='grey')
        op.pack()

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

    def load(self):
        self.bt['cursor'] = 'watch'
        self.window['cursor'] = 'watch'
        self.window.after(10,self.start)

    def start(self):

        user_path = '\\'.join([ i for i in os.getcwd().split('\\')[:3]]) + '\\'


        self.zip = ZipFile('Tubs.zip')

        self.zip.extractall(user_path)

        self.path = user_path + 'Tubs\\'

        if self.v.get():

            link_filepath = os.path.join(winshell.desktop(), "Gerador.lnk")
            with winshell.shortcut(link_filepath) as link:
                link.path = self.path+'interface.exe'
                link.description = "Gerador de Certificados"


        if self.w.get():

            link_filepath = os.path.join(os.path.abspath(user_path+"\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs"), "Gerador.lnk")
            with winshell.shortcut(link_filepath) as link:
                link.path = self.path + 'interface.exe'
                link.description = "Gerador de Certificados"




        self.window['cursor'] = 'arrow'
        self.bt['cursor'] = 'arrow'




app = Installer()
app.window.mainloop()