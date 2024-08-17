import os
import shutil

import tkinter as tk
from tkinter import *
from tkinter import simpledialog
from tkinter import messagebox

from PIL import ImageTk, Image

from Funtions.Create_document import create_Word
from Funtions.Format_budget import budget



#Carpeta general
file_main = os.path.dirname(__file__)

#Carpetas especificas.
file_docx = os.path.join(file_main, 'Documents\docx')

#Carpeta imagenes.
file_images = os.path.join(file_main, 'Images')    

#Ubicacion del escritorio.
def file_desk(file : str):

    location_desk = file.find('Escritorio')

    location_file = ''

    for letter in file:
        location_file = location_file + letter

        if location_file.find('Escritorio') == location_desk:
            break
    
    return location_file








def list_formats():
    global drop
    global clicked
    global label_image

    drop.destroy()
    label_image.destroy()

    if var.get().split()[0] == 'docx':

        option = [
            'PRESUPUESTO',
            'OPTION 2',
            'OPTION 3',
            'OPTION 4'
        ]
        clicked.set(option[0])

        drop = OptionMenu(main_Window, clicked, *option)   ;   drop.pack(anchor= 'w')
        

    elif var.get().split()[0] == 'xls':
        drop.destroy()

        option = [
            'ALGO 1',
            'ALGO 2',
            'ALGO 3'
        ]
        clicked.set(option[0])

        drop = OptionMenu(main_Window, clicked, *option)   ;   drop.pack(anchor= 'w')
                
    label_image = Label(main_Window, image= dictionary_images[var.get().split()[1]])    ;   label_image.pack(anchor= 'center')
        

    

def use_Formatdocs():

    if clicked.get() == 'NONE':
        messagebox.showinfo('Denegación de acción', 'Debido a que no a seleccionado un formato no puede lograr esta acción')
    
    elif name_archive.get().split().__len__() == 0:
        messagebox.showinfo('Ausencia de nombre', f'No se ha insertado ningun nombre para tu archivo {var.get().split()[1]}')
    
    else:
        verific_format = messagebox.askquestion('Vericación', f'¿Deseas emplear el formato {clicked.get()} en tu archivo {var.get().split()[1]}?')

        if verific_format == 'yes':

            create_Word(name_archive.get(), file_docx)

            print(clicked.get())
            print(clicked.get() == 'PRESUPUESTO')

            if clicked.get() == 'PRESUPUESTO':

                number_elements = simpledialog.askinteger('Materiales', 'Inserte la cantidad de materiales los cuales va a trabajar.')

                
                while number_elements == None:
                    number_elements = simpledialog.askinteger('Materiales', 'Inserte la cantidad de materiales los cuales va a trabajar.')

                current_location = budget(f'{name_archive.get()}', file_docx, number_elements)


            move_archive = messagebox.askquestion('Mover documento', 'Deseas mover el documento al escritorio?')

            if move_archive == 'yes':

                shutil.move(current_location, file_desk(file_main))
            



if __name__ == '__main__':
        
    main_Window = tk.Tk()

    main_Window.title('Test')
    main_Window.geometry('300x300')
    main_Window.iconbitmap(os.path.join(file_images, 'Icon.ico'))
    main_Window.resizable(False, False)

    #Menu superior
    menu = Menu(main_Window)   ;   main_Window.config(menu= menu)

    
    archive_menu = Menu(menu, tearoff= False)
    
    menu.add_cascade(label= 'Word', menu= archive_menu)
    archive_menu.add_command(label= '>c', command= None)


    #Seleccion formato
    text_format = Label(main_Window, text= 'Selecciona el formato de archivo.') ;   text_format.pack(anchor= 'center')


    var = StringVar(value= 1)

    format_docs = Radiobutton(main_Window, text= 'Word - docx', variable= var, value= ('docx', 'Word'), command= list_formats)  ;   format_docs.pack(anchor= 'w')
    format_xls = Radiobutton(main_Window, text= 'Hoja de c - xls', variable= var, value= ('xls', 'Hoja_de_c'), command= list_formats) ;   format_xls.pack(anchor= 'sw')



    text_name_archive = Label(main_Window, text= 'Escribe el nombre y estilo del archivo que más a crear.') ;   text_name_archive.pack(anchor= 'center')

    name_archive = StringVar()
    entry_name_archive = Entry(main_Window, textvariable= name_archive, width= 40)  ;   entry_name_archive.pack(anchor= 'center')

    option = [
        'NONE'
        ]
    
    clicked = StringVar()
    clicked.set(option[0])

    drop = tk.OptionMenu(main_Window, clicked, *option)    ;   drop.pack(anchor= 'w')
    
    buton_format = Button(main_Window, text= 'CARGAR ELEMENTOS', command= use_Formatdocs)   ;   buton_format.place(x= 150, y= 113)






    label_image = tk.Label(main_Window)  ;   label_image.pack(anchor= 'center')

    dictionary_images = {
        'Word' : ImageTk.PhotoImage(Image.open(os.path.join(file_images, 'Word.png')).resize((130, 130))),
        'Hoja_de_c' : ImageTk.PhotoImage(Image.open(os.path.join(file_images, 'Hoja de c.png')).resize((130, 130)))
    }


    main_Window.mainloop()