
'''instale o python e o pacote abaixo'''
#pip install pdf2docx
from pdf2docx import Converter
# convert pdf to docx
from tkinter import *
from tkinter import filedialog as fd
from tkinter import ttk 
app = Tk()
import time
app.title("Conversor de PDF para Word")
app.geometry('850x300')
var = StringVar()
#progress bar
valor = DoubleVar()
progress = ttk.Progressbar(app,variable=valor, maximum=100) 
progress.pack(pady = 10)
valor.set(0)
def bar():
    a=10
    while a < 110:
        print(a)
        valor.set(a)
        time.sleep(0.1)
        a+=1
        app.update()
def converter():
    pdf_file= fd.askopenfilename() 
    if pdf_file:
        label.config(text='Aguarde...!',fg='red',font=("Courier", 30))
        bar()#funcao progress bar
    docx_file =(tela.get()+'.docx')
    try:
        if docx_file == '.docx':
            label.config(text='coloque um nome abaixo e\nclique no botão Adicionar PDF!',fg='red',font=("Courier", 20))
        else:
            cv = Converter(pdf_file)
            cv.convert(docx_file, start=0, end=None)
            cv.close()
            label.config(text=docx_file+' convertido com sucesso!',fg='blue',font=("Courier", 15))
            progress.pack_forget()
    except:
        label.config(fg='red',text='erro ao converter arquivo\n verifique se o nome está correto!\n ou se o aquivo está na mesma pastas!')
label1 = Label(app, text='Conversor de PDF para Word')
label1.config(font=("Courier", 20))
label1.pack()

var = 'De um nome para o arquivo que será convertido \n no campo abaixo antes de clicar no botão'
label = Label(app, text=var)
label.config(font=("Courier", 20))
label.pack()
tela = Entry(app,font=("Calibri 22"))
tela.pack()
bt1 = Button(app, text='Adicionar PDF', bg='green', fg='white',font=("Calibri 22"),command=converter)
bt1.pack()
app.mainloop()
'''
from pdf2docx import Converter
pdf_file = 'oi.pdf'
docx_file = 'jorge.docx'

# convert pdf to docx
try:
    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)
    cv.close()
except:
    print('erro')
'''
