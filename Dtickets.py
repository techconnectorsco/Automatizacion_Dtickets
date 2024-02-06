#Archivos importados necesarios para la ventana inicial
from os import close
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QProgressDialog, QVBoxLayout,QMessageBox, QPushButton, QFileDialog, QLabel, QComboBox, QHBoxLayout, QLineEdit
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt , QTimer

#Archivos necesarios para la automatizacion en selenium
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import json
import pandas as pd
import re
import generador_pdf
import requests

"""
formato correcto de las columnas en el archivo de excel para tomar en cuenta es :
    'EVENTO', 'ASISTENTE', 'EMAIL', 'CEDULA', 'TELEFONO', 'CANTIDAD', 'TIPO'
    """


pd.set_option('display.max_columns', None)

df_enviados = pd.DataFrame(columns=['EVENTO', 'ASISTENTE', 'EMAIL', 'CEDULA', 'TELEFONO', 'CANTIDAD', 'TIPO'])
evento_no_disponible = []
def mostrar_mensaje_error(mensaje):
    """ funcion que muestra una ventana con el mensaje del error si el archivo 
        no se encuentra en la ruta especificada"""
    
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText("Error")
    msg.setInformativeText(mensaje)
    msg.setWindowTitle("Error")
    msg.exec_()


def app_window():
    """Ventana Inicial que abre la aplicacion para determinar la ruta de descarga de la data de cortesia
        por el momento se genero un servidor local para simular la forma en la que se descarga el archivo
        
        ruta : http://127.0.0.1:5000
        endpoint : Data_Cortesias(eventos en un solo archivo) o endpoint: Cortesia_Test(Nombre del evento especifico)
        """

    url_local = "http://127.0.0.1:5000/Data_Cortesias"
    class Ventana(QWidget):

        def __init__(self):
            super().__init__()

            self.setWindowTitle('Descargar Data')
            self.setGeometry(100, 100, 400, 200)

            layout = QVBoxLayout()
            self.setLayout(layout)

            self.dticon = QPixmap('images/dticon.png')
            self.logo = QPixmap('images/Dtickets.png')
            self.setWindowIcon(QIcon(self.dticon))

            self.imagen = QLabel(self)
            self.imagen.setPixmap(self.logo.scaled(300, 300, Qt.KeepAspectRatio))
            layout.addWidget(self.imagen)
            self.imagen.setAlignment(Qt.AlignCenter)

            label_evento = QLabel("Seleccione Evento", self)
            layout.addWidget(label_evento)
            label_evento.setAlignment(Qt.AlignCenter)

            self.label = QLabel("Data de Evento a Descargar")
            layout.addWidget(self.label)
            label_font = self.label.font()
            label_font.setBold(True)
            label_font.setPointSize(10)
            self.label.setAlignment(Qt.AlignCenter)

            self.url_input = QLineEdit(url_local, self)
            font = self.url_input.font()
            font.setBold(True)
            font.setPointSize(12)
            self.url_input.setFont(font)
            self.url_input.setAlignment(Qt.AlignCenter)
            self.url_input.setReadOnly(False)
            layout.addWidget(self.url_input)

            self.boton = QPushButton('Comenzar Automatización')
            self.boton.clicked.connect(self.cargar_archivo)
            self.boton.setStyleSheet("""
                QPushButton {
                    background-color: #22367D;
                    color: #D7E749;
                    border: 1px solid #D7E749;
                    padding: 10px;
                    min-width: 50px;
                    min-height: 20px;
                }
                QPushButton:hover {
                    background-color: #3D3EB4;
                }
            """)
            layout.addWidget(self.boton)

        def cargar_archivo(self):
            """Carga del archivo descargado de la web"""

            url = self.url_input.text()
            nombre_archivo = "Data.xlsx"

            if not url:
                mostrar_mensaje_error("Por favor, ingrese una URL válida.")
                return

            self.descargar_archivo(url, nombre_archivo)
            #Continúa con el resto del código después de la descarga

        def descargar_archivo(self, url, nombre_archivo):
            try:
                Data_xlsx = requests.get(url)
                Data_xlsx.raise_for_status()
                with open(nombre_archivo, 'wb') as archivo:
                    archivo.write(Data_xlsx.content)
                    time.sleep(2)
                    print("Archivo Descargado Exitosamente")
                
                self.close()  # Cerrar la ventana después de la descarga exitosa
                selenium_code(nombre_archivo)
            except requests.exceptions.RequestException as e:
                mostrar_mensaje_error(f"Error al descargar el archivo: {e}")

    app = QApplication([])
    ventana = Ventana()
    ventana.show()
    app.exec_()
    

class SeleniumAutomation:
    def __init__(self):
        self.browser = self.setup_browser()

    def setup_browser(self):

        """configura y lanza un navegador Chrome con ciertas opciones. Ignora los errores de certificado y SSL, 
        desactiva QUIC y maximiza la ventana del navegador. Devuelve el objeto del navegador para su uso posterior."""

        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument("disable-quic")
        options.add_argument("--start-maximized")
        options.add_argument('--disable-proxy-certificate-handler')
        return webdriver.Chrome(options=options)

    def navigate_to_site(self, url):

        """Esta función toma como argumento la URL específica y hace clic en un elemento de la página. 
        Luego, espera 3 segundos para permitir que la página se cargue completamente."""

        self.browser.get(url)
        WebDriverWait(self.browser,8).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div/div[4]/div/div/div/a"))).click() #click al menu inicial
        time.sleep(3)

    def login_to_site(self, credentials_path):

        """Hace clic en un elemento para iniciar el proceso de inicio de sesión,lee las credenciales de un archivo JSON, 
        introduce las credenciales en los campos correspondientes de la página web y hace clic en el botón de inicio 
        de sesión. Luego, espera 10 segundos para permitir que el proceso de inicio de sesión se complete."""
        

        WebDriverWait(self.browser,6).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='elementor-popup-modal-451']/div/div[2]/div/div/div/div/div[4]/div/div[1]/div/div/div/div[1]/a/span[2]"))).click() #click a iniciar sesion
        with open(credentials_path) as f:
            credentials = json.load(f)
        time.sleep(3)
        self.browser.execute_script("window.scrollBy(0, 220);")
        username = WebDriverWait(self.browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='user']"))) #campo de usuario
        username.clear()
        username.send_keys(credentials['username'])
        password = WebDriverWait(self.browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='password']"))) #campo de contraseña
        password.clear()
        password.send_keys(credentials['password'])
        WebDriverWait(self.browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='content']/div/div[1]/div/div/div[2]/div/form/div/div[4]/button"))).click() #boton para iniciar sesion

        
    def find_events(self, nombre_archivo):

        """Busca los eventos en la página web.

        Esta función espera hasta que un elemento específico en la página web esten disponibles
        Luego, encuentra todos los elementos hijos de ese elemento que tienen una clase específica.
        Itera sobre estos elementos hijos y los almacena para un posteror proceso.


        Parámetros:
        nombre del archivo a tratar

        Devoluciones:
        Ninguna"""

        data_cortesia, data_erronea = SeleniumAutomation.read_excel_data(self, nombre_archivo)

        #ubica el div padre contenedor para luego ubicar cada div hijo de cada evento de forma individual y los cuenta
        eventos = WebDriverWait(self.browser, 8).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/section/div/div[2]/div/div[2]/div/div")))
        divs_evento = eventos.find_elements(By.CLASS_NAME, "etn-col-lg-7")
        botones = dict()
        coincidencias_df = pd.DataFrame(columns=['EVENTO', 'ASISTENTE', 'EMAIL', 'CEDULA', 'TELEFONO', 'CANTIDAD', 'TIPO'])
        no_encontrado = set()
        # Iterar sobre cada fila del DataFrame para comparar con el texto del evento actual
        for index, row in data_cortesia.iterrows():
            
            for i in range(1, len(divs_evento)+1):
                buscar_evento = WebDriverWait(self.browser,8).until(EC.element_to_be_clickable((By.XPATH, f"/html/body/div[2]/section/div/div[2]/div/div[2]/div/div/div[{i}]/div/div/div/div/h3/a" )))
                if SeleniumAutomation.limpiar_cadena(row['EVENTO']) in SeleniumAutomation.limpiar_cadena(buscar_evento.text):
                    botones[SeleniumAutomation.limpiar_cadena(row['EVENTO'])] = f"/html/body/div[2]/section/div/div[2]/div/div[2]/div/div/div[{i}]/div/div/div/div/h3/a"
                    coincidencias_df = coincidencias_df._append({
                    'EVENTO': SeleniumAutomation.limpiar_cadena(row['EVENTO']),
                    'ASISTENTE': row['ASISTENTE'],
                    'EMAIL': row['EMAIL'],
                    'CEDULA': row['CEDULA'],
                    'TELEFONO': row['TELEFONO'],
                    'CANTIDAD': row['CANTIDAD'],
                    'TIPO': row['TIPO']
                }, ignore_index=True)
                    
                else:
                    no_encontrado.add(SeleniumAutomation.limpiar_cadena(row['EVENTO']))          

        pd.set_option('display.float_format', '{:.0f}'.format)
        

        for evento, xpath in botones.items():
            boton = WebDriverWait(self.browser, 8).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            boton.click()
            time.sleep(2)
            url_cortesia_evento = self.browser.current_url
            SeleniumAutomation.iteracion_de_cada_evento(self.browser, coincidencias_df, evento, url_cortesia_evento)
        
        
        generador_pdf.generate_pdf(df_enviados, data_erronea, data_cortesia) #eventos_encontrados, evento_sin_encontrar coincidencias_df
        time.sleep(2)
        self.browser.close()
        self.browser.quit()

    def iteracion_de_cada_evento(browser,df_cortesia, evento, url_cortesia_evento):
        global df_enviados
        pd.set_option('display.float_format', '{:.0f}'.format)
        grupos_por_evento = df_cortesia.groupby('EVENTO')

        for nombre_grupo, grupo in grupos_por_evento:
            if nombre_grupo == evento:
                try:
                    asistentes_evento = grupo[['EVENTO', 'ASISTENTE', 'EMAIL', 'CEDULA', 'TELEFONO', 'CANTIDAD', 'TIPO']]
                    # Intentar encontrar el elemento
                    elemento = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div/div[2]/div/div[2]/div/form"))) #formulario de cortesias
                    if elemento:
                        time.sleep(1)
                        for index, asistente in asistentes_evento.iterrows():
                            time.sleep(2)
                            try:
                                alert = browser.switch_to.alert
                                alert.accept()  # O alert.dismiss() según sea necesario
                            except:
                                pass
                            browser.execute_script("window.scrollBy(0, 280);")
                            cantidad = WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ticket-input_0']"))) #campo cantidad
                            cantidad.clear()
                            cantidad.send_keys(1)
                            WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div/div[2]/div/div[2]/div/form/input[7]"))).click() #click al boton buy ticket        
                            time.sleep(1.5)
                            #url_actual = browser.current_url
                            nombre = WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ticket_0_attendee_name_1']")))
                            nombre.clear()
                            nombre.send_keys(asistente['ASISTENTE'])

                            correo = WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ticket_0_attendee_email_1']")))
                            correo.clear()
                            correo.send_keys(asistente['EMAIL'])

                            cedula = WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='etn_attendee_extra_field_0_attendee_1_input_0']")))
                            cedula.clear()
                            cedula.send_keys(int(asistente['CEDULA']))
                            time.sleep(2)

                            WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='etn-event-attendee-data-form']/div[3]/button"))).click()
                            nombre_completo = asistente['ASISTENTE']
                            nombre, apellido = nombre_completo.split(' ', 1)

                            time.sleep(1)

                            nombre_factura = WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='billing_first_name']")))
                            nombre_factura.clear()
                            nombre_factura.send_keys(nombre)

                            time.sleep(1)

                            apellido_factura = WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='billing_last_name']")))
                            apellido_factura.clear()
                            apellido_factura.send_keys(apellido)

                            time.sleep(1)

                            telefono_factura = WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='billing_phone']")))
                            telefono_factura.clear()
                            telefono_factura.send_keys(asistente['TELEFONO'])

                            time.sleep(1)
                            telefono_factura = WebDriverWait(browser, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='billing_email']")))
                            telefono_factura.clear()
                            telefono_factura.send_keys(asistente['EMAIL'])

                            time.sleep(1)

                            browser.execute_script("window.scrollBy(0, 600);")

                            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='terms']"))).click()

                            time.sleep(1)
                                        
                            WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='place_order']"))).click()

                            time.sleep(5)
                            
                            #WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='content']/div/div[1]/div/div/div[2]/div/div/div/section[1]/table/tbody/tr/td[1]/a"))).click()
                            browser.get(url_cortesia_evento)

                        df_enviados = pd.concat([df_enviados, asistentes_evento], ignore_index=True)
                        browser.get('https://wordpressmu-957299-4051864.cloudwaysapps.com/panel-de-control/')

                    else:
                        
                        browser.get('https://wordpressmu-957299-4051864.cloudwaysapps.com/panel-de-control/')

                except TimeoutException:
                    browser.get('https://wordpressmu-957299-4051864.cloudwaysapps.com/panel-de-control/')

                except NoSuchElementException:
                    browser.get('https://wordpressmu-957299-4051864.cloudwaysapps.com/panel-de-control/')

        


    def limpiar_cadena(cadena):
        """Funcion que normaliza la cadena mas especificamente los nombres de los eventos 
            para facilitar el tratamiento de los datos"""
        cadena = re.sub(r'[^a-zA-Z0-9\s]', '', cadena)  # Eliminar caracteres especiales
        cadena = cadena.replace(' ', '').replace('-', '').replace('–','').replace('—','').replace('í','i')
        cadena = cadena.strip().lower()
        return cadena


    def read_excel_data(self, excel_path):
        """
        Lee un archivo Excel y valida las direcciones de correo electrónico en él.

        Esta función lee un archivo Excel especificado por el parámetro `excel_path`, 
        itera sobre sus filas y valida las direcciones de correo electrónico encontradas en la tercera columna de cada fila.
        Las direcciones de correo electrónico se validan contra un patrón de expresión regular.
        Si una dirección de correo electrónico no coincide con el patrón, se considera incorrecta.
        Las direcciones de correo electrónico incorrectas se almacenan en el diccionario `self.email_incorrecto` con el nombre asociado.

        Parámetros:
        excel_path (str): La ruta al archivo Excel a leer.

        Devoluciones:
        devuelve un Dataframe con la data correctas para su tratamiento en la automatizacion
         y data erronea que no entrara en la automatizacion """

        df = pd.read_excel(excel_path, dtype={'TELEFONO': str})
        df_correcto = pd.DataFrame(columns=df.columns)  # DataFrame vacío con las mismas columnas que df
        df_incorrecto = pd.DataFrame(columns=df.columns)
        #print(df)


        for index, row in df.iterrows():
        # Validación de correo electrónico
            pattern_email = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
            is_email_valid = re.match(pattern_email, row.iloc[2].lower())

            # Validación de número de cédula
            cedula = int(row.iloc[3])  # Asegúrate de que la cédula sea tratada como cadena
            is_cedula_valid_int = (
                cedula > 0
            )
            cedula = str(cedula)
            is_cedula_valid_len = (
            8 <= len(cedula) <= 10
            and cedula.isdigit()  # Verifica si todos los caracteres son dígitos
        )
            
            # Agregar al DataFrame correspondiente según las validaciones
            if is_email_valid and is_cedula_valid_len and is_cedula_valid_int:
                df_correcto.loc[index] = row
            else:
                df_incorrecto.loc[index] = row
        
        
        return df_correcto, df_incorrecto

def selenium_code(nombre_archivo):
    automation = SeleniumAutomation()
    automation.navigate_to_site("https://wordpressmu-957299-4051864.cloudwaysapps.com/")
    automation.login_to_site('credentials.json')
    automation.find_events(nombre_archivo)
    
if __name__=='__main__':
    
    app_window()