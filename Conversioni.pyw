# -*- coding: utf-8 -*-
"""
Created on Fri Feb  5 14:18:35 2021

@author: Eleonora
"""
import os
import PySimpleGUI as sg
from pdf2image import convert_from_path
from PIL import Image
from pdf2docx import parse


sg.theme("DarkTeal2")
layout = [[sg.Text("Se vuoi trasformare un pdf in immagine/i:"), sg.Button("pdf2jpg")],[sg.Text("Se vuoi trasformare una o più immagini in un pdf:"), sg.Button("jpg2pdf")], [sg.Text("Se vuoi trasformare un pdf in un word:"), sg.Button("pdf2word")]]
#layout = [[sg.Button("pdf2jpg")],[sg.Button("jpg2pdf")],[sg.Button("pdf2word")]]

###Building Window
window = sg.Window('Conversione', layout, size=(500,200))
    
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Esci":
        break
    elif event == "pdf2jpg":
        # prima opzione: da pdf a immagine/i
        pdf_name = sg.popup_get_file('Scegli il pdf')
        try : 
            pp = os.path.dirname(os.path.abspath(pdf_name)) # percorso della cartella
            images = convert_from_path(pdf_name)
            for i in range(len(images)):	
            # Save pages as images in the pdf
                images[i].save(pp +'\Pag'+ str(i) +'.jpg', 'JPEG')
            sg.popup('Finito di convertire il PDF')
        except:
            sg.popup('Ops, qualcosa è andato storto, clicca OK e riprova')
            pass
        
    elif event == "jpg2pdf":
        n_jpg = sg.popup_get_text("Quante immagini da trasformare in unico pdf?")
        try:
            nn = int(n_jpg)
            for ii in range(nn):
                jpg_name = sg.popup_get_file("Scegli l immagine numero " + str(ii+1) + ":")
                if ii == 0:
                    image0 = Image.open(jpg_name)
                    pp = os.path.dirname(os.path.abspath(jpg_name)) # percorso della cartella
                    im0 = image0.convert('RGB')
                    imagelist = []
                else:
                    image1 = Image.open(jpg_name)
                    im1 = image1.convert('RGB')
                    imagelist.append(im1)
            im0.save(pp + '\\nuovoPDF.pdf', save_all ='TRUE', append_images = imagelist)
            sg.popup('Finito di creare il PDF')
                    
        except:
            sg.popup('Ops, qualcosa è andato storto, clicca OK e riprova')
            pass
        
    elif event == "pdf2word":
        word_name = sg.popup_get_file('Scegli il pdf') # in realtà è un pdf
        pp = os.path.dirname(os.path.abspath(word_name)) # percorso della cartella
        try:
            parse(word_name, pp + '\\Nuovoword.docx',start = 0, end = None)
            sg.popup('Finito di creare il word')
        except:
            sg.popup('Ops, qualcosa è andato storto, clicca OK e riprova')
            pass
                
            

            
            
            
            
        
        
        
        
        
        
        