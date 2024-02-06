from typing import Self
from fpdf import FPDF
from datetime import datetime
import os
import re
from numpy import int64
import pandas as pd
pd.set_option('display.float_format', '{:.0f}'.format)


import Dtickets

def corregir_palabras(texto):
    # Dividir palabras utilizando expresiones regulares
    if '@' in texto:
         return texto
    else:
        palabras = re.findall(r'\b\w+\b', texto)
        
        # Corregir capitalización y unir las palabras
        texto_corregido = ' '.join(palabra.capitalize() for palabra in palabras)
        
        return texto_corregido

def generate_pdf(data_cortesia, data_erronea, data_correcta):#, eventos_en_la_web, eventos_en_data):
    ahora = datetime.now()
    fecha = ahora.strftime('%d-%m-%Y')

    class PDF(FPDF):
        def header(self):
            self.image('images/fondo.png', x=6, y=5, w=199, h=23)
            self.set_font('Arial', 'B', 12)
            pdf.set_draw_color(34, 54, 125)
            pdf.rect(x=6, y=6, w=199, h=286)
            pdf.add_font(family= 'Raleway_Bold', style='', fname='fonts/Raleway-SemiBold.ttf', uni=True)
            pdf.add_font(family= 'Raleway', style='', fname='fonts/Raleway-Light.ttf', uni=True)
            self.image('images/Dtickets.png', 7,8,38,18)
            self.image('images/python.png', 183,8,8,8)
            self.image('images/selenium.png', 195,8,8,8)
            self.set_font('Raleway_Bold','',20)
            self.set_text_color(r= 215 , g= 231, b= 73 )
            self.multi_cell(w=0,h=1,txt='', border=0, align='C', fill=0,)
            self.multi_cell(w=0,h=12,txt=" CORTESIAS D'TICKETS ", border=0, align='C', fill=0,)
            self.multi_cell(w=0,h=8,txt='', border=0, align='C', fill=0,)
            
            fecha_realizado = ahora.strftime('%d/%m/%Y %H:%M:%S')
            pdf.set_font('Raleway','',10)
            pdf.text(170,22,'Elaborado el:')
            pdf.text(170,27,fecha_realizado)

        def footer(self):

            pdf.add_font(family= 'Raleway_Bold', style='', fname='fonts/Raleway-SemiBold.ttf', uni=True)
            pdf.add_font(family= 'Raleway', style='', fname='fonts/Raleway-Light.ttf', uni=True)            
            pdf.set_font('Raleway','',6)
            self.image('images/tech_logo.png', 18,270,15,15)
            pdf.text(8,286,'AUTOMATIZACIóN REALIZADA POR EL EQUIPO DE TECHCONNECTORS.CO')
            pdf.set_font('Raleway_Bold','',8)
            pdf.text(8,290,'Devs@techconnectors.co')
            pdf.text(83,290,'Generado con Python y Selenium')
            self.set_y(-10)
            
            self.set_font('Raleway_Bold', '', 8)
            
            self.cell(w=0, h=6, txt='Pagina ' + str(self.page_no()) + '/{nb}', border=0, align='R', fill=0)



    pdf = PDF()

    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.add_font(family= 'Raleway_Bold', style='', fname='fonts/Raleway-SemiBold.ttf', uni=True)
    pdf.add_font(family= 'Raleway', style='', fname='fonts/Raleway-Light.ttf', uni=True)


    pdf.set_top_margin(10)
    pdf.set_left_margin(8)
    pdf.set_right_margin(8)

    pdf.set_font('Raleway_Bold', '', 12)
    pdf.cell(0, 8, 'CORTESIAS ENVIADAS EXITOSAMENTE', 0, 1, 'C')

    if data_cortesia.empty:
        pdf.set_font('Raleway', '', 14)
        pdf.cell(0, 8, 'No se logro enviar ninguna cortesia', 0, 1, 'C')
        pdf.cell(0, 10, '', 0, 1, 'C')
    else:
        pd.set_option('display.float_format', '{:.0f}'.format)
        data_cortesia['CEDULA'] = data_cortesia['CEDULA'].astype(int)
        columnas_seleccionadas = ['EVENTO', 'ASISTENTE', 'EMAIL', 'CEDULA']
        datos_seleccionados = data_cortesia[columnas_seleccionadas]
        columnas = datos_seleccionados.columns.tolist()
        datos = datos_seleccionados.values.tolist()

        pdf.set_font("Raleway_Bold", size = 11)
        pdf.set_draw_color(34, 54, 125)
        pdf.set_line_width(0.3)
        # Crear las cabeceras de la tabla
        for i in range(len(columnas)):
            pdf.cell(w=48,h=6, txt = columnas[i], border = 1, align = 'C',)
        pdf.ln()

        # Crear las filas de la tabla
        pdf.set_font("Raleway", size = 10)  
        for i in range(len(datos)):
            for j in range(len(datos[i])):
                texto = str(datos[i][j])
                texto = texto.replace(u'\u2013', '-')
                pdf.cell(w=48,h=6, txt = corregir_palabras(texto), border = 0.5, align = 'C',)
            pdf.ln()
    
    pdf.cell(0, 10, '', 0, 1, 'C')
    if data_erronea.empty:
         pass
    else:
        pd.set_option('display.float_format', '{:.0f}'.format)
        data_erronea['CEDULA'] = data_erronea['CEDULA'].astype(int)
        pdf.set_font('Raleway_Bold', '', 12)
        #pdf.cell(0, 8, '', 0, 1, 'C')
        pdf.cell(0, 8, 'CORTESIAS NO ENVIADAS POR ERROR EN LOS DATOS', 0, 1, 'C')

        columnas_seleccionadas = ['EVENTO', 'ASISTENTE', 'EMAIL', 'CEDULA']
        datos_con_error = data_erronea[columnas_seleccionadas]
        columnas = datos_con_error.columns.tolist()
        datos_con_error_de_email = datos_con_error.values.tolist()

        pdf.set_font("Raleway_Bold", size = 11)
        pdf.set_draw_color(34, 54, 125)
        pdf.set_line_width(0.3)
        # Crear las cabeceras de la tabla
        for i in range(len(columnas)):
            pdf.cell(w=48,h=6, txt = columnas[i], border = 1, align = 'C',)
        pdf.ln()

        # Crear las filas de la tabla
        pdf.set_font("Raleway", size = 10)  
        for i in range(len(datos_con_error_de_email)):
            for j in range(len(datos_con_error_de_email[i])):
                texto = str(datos_con_error_de_email[i][j])
                texto = texto.replace(u'\u2013', '-')
                pdf.cell(w=48,h=6, txt = texto, border = 0.5, align = 'C',)
            pdf.ln()

    pdf.cell(0, 10, '', 0, 1, 'C')
    if data_correcta.empty:
        pass
    else:
        pd.set_option('display.float_format', '{:.0f}'.format)
        data_correcta['CEDULA'] = data_correcta['CEDULA'].astype(int64)
        pdf.set_font('Raleway_Bold', '', 12)
        pdf.cell(0, 8, 'DATOS PROCESADOS CORRECTAMENTE', 0, 1, 'C')

        columnas_seleccionadas = ['EVENTO', 'ASISTENTE', 'EMAIL', 'CEDULA']
        datos_sin_error = data_correcta[columnas_seleccionadas]
        columnas = datos_con_error.columns.tolist()
        datos_sin_error_de_email = datos_sin_error.values.tolist()

        pdf.set_font("Raleway_Bold", size = 11)
        pdf.set_draw_color(34, 54, 125)
        pdf.set_line_width(0.3)
        # Crear las cabeceras de la tabla
        for i in range(len(columnas)):
            pdf.cell(w=48,h=6, txt = columnas[i], border = 1, align = 'C',)
        pdf.ln()

        # Crear las filas de la tabla
        pdf.set_font("Raleway", size = 10)  
        for i in range(len(datos_sin_error_de_email)):
            for j in range(len(datos_sin_error_de_email[i])):
                texto = str(datos_sin_error_de_email[i][j])
                texto = texto.replace(u'\u2013', '-')
                pdf.cell(w=48,h=6, txt = texto, border = 0.5, align = 'C',)
            pdf.ln()
    

    pdf.output(fecha +'.pdf', 'F')
    path = fecha +'.pdf'
    os.system(path)