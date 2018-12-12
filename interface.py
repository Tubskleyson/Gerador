from tkinter import filedialog
from tkinter import *
from PIL import ImageTk, Image

class App:

    def __init__(self):

        self.window = Tk()

        self.window.geometry('800x500')

        self.window.configure(background='white')



        space = Label(self.window,bg='white')
        space.pack(pady=20)


        load = Image.open("cert.png")

        load = load.resize((200, 200), Image.ANTIALIAS)

        img = ImageTk.PhotoImage(load)



        image = Label(self.window,text = 'kalfoi', image = img, background='white')
        image.image = img
        image.pack()

        tit = Label(self.window,text='Gerador de Certificados', bg='white',font="sans 25 ", fg='grey')
        tit.pack()

        bt = Button(self.window,text='Novo Pacote', height=2, width=13, cursor='hand2', command=self.select)
        bt.pack(pady=10)

        self.widgets = [space,image,tit,bt]

    def select(self):

        for i in self.widgets: i.destroy()

        tit = Label(self.window, text='Selecionar Documentos', bg='white', font="sans 25 ", fg='grey')
        tit.pack(pady=30)

        left = Frame(bg='white')
        left.pack(side=LEFT, fill=Y,pady=50)

        lp = Image.open("cert.jpg")

        lp = lp.resize((100, 100), Image.ANTIALIAS)

        li = ImageTk.PhotoImage(lp)

        imagel = Label(left, text='kalfoi', image=li, background='white')
        imagel.image = li
        imagel.pack(side=TOP)


        b0 = Button(left,text='Selecionar Modelo (.docx)', width=25,height=2, command=lambda : self.get_file(0))
        b0.pack(padx=100, pady=20)

        right = Frame(bg='white')
        right.pack(side=LEFT, fill=Y, pady=50)

        rp = Image.open("list.png")

        rp = rp.resize((100, 100), Image.ANTIALIAS)

        ri = ImageTk.PhotoImage(rp)

        imager = Label(right, text='kalfoi', image=ri, background='white')
        imager.image = ri
        imager.pack(side=TOP)

        b1 = Button(right,text='Selecionar Lista (.csv)', width=25,height=2, command=lambda : self.get_file(1))
        b1.pack(padx=100, pady=20)


        print()

    def get_file(self,x):

        if not x: return filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("Documento Word", "*.docx"), ("all files", "*.*")))
        else: return filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("Planilha", "*.csv"), ("all files", "*.*")))





app = App()
app.window.mainloop()