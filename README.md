# Automatización de Descargas y Registro de Eventos para D'Tickets

Este script de Python utiliza la biblioteca Selenium para automatizar la descarga y registro de eventos en un sitio web basado en WordPress. 
El script incluye una interfaz gráfica (GUI) construida con PyQt5 para facilitar la interacción del usuario.

## Requisitos

Asegúrate de tener Python y las siguientes bibliotecas instaladas:

- PyQt5
- Selenium
- pandas
- requests
Para mejor manejo de las bibliotecas de este proyecto instalar los Requirements.txt

# Uso
- Ejecuta el script y se abrirá una ventana para ingresar la URL del servidor local y seleccionar el evento.

- Haz clic en "Comenzar Automatización" para iniciar el proceso de descarga y registro.

# Configuración

La URL del servidor local de pruebeas predeterminada es  "http://127.0.0.1:5000/Data_Cortesias", pero puedes cambiarla según tus necesidades para apuntar al endpoint correspondiete
para la descarga del archivo de excel.

Asegúrate de tener el archivo de credenciales credentials.json con las credenciales necesarias para iniciar sesión en el sitio web
## Editar el archivo credentials.json para agregar el usuario y contraseña correspondiente

# El formato correcto para el archivo de Excel es:

'EVENTO', 'ASISTENTE', 'EMAIL', 'CEDULA', 'TELEFONO', 'CANTIDAD', 'TIPO'
como cabecera o titulos de cada columna del archivo en Excel
