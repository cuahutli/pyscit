#! /usr/bin/python
# -*- coding: utf-8 -*-

"""
    Modulo creado por: Cuahutli Miguel Ulloa Robles
    Version: 1.0
"""

"""
    Clase para enviar correo electronico, se necesitan pasar los parametros servidor,puerto,usuario,pass
    receptores,asunto,mensaje y los archivos adjuntos.

    se pueden modificar  los que vienen por default para pasar menos parámetros si el servidor y emisor de
    los correos siempre es el mismo puedes ponerlos como por default para no pasarlos siempre:

    En mi caso modificaré por los parámetros del mail de SAMAO y sólo pasaré Receptores, asunto, mensaje 
    y adjutos haciendo más rápido el proceso desde mis clases
"""

"""
    para utilizarlo solo hay que colocar el .py en la ruta de los archivos que van a utilizarlo
    instanciar la clase y llamar a la funcion pasandole los parametros, ejemplo:

        from correo import manda_mail
        if __name__== "__main__":
                correo = manda_mail() #creando instancia
                receptores =['correo@dominio.com', 'correo2@dominio.com']
                asunto = "esto Funciona"
                mensaje = "mensaje <b> xD </b>"
                adjuntos = ['archivo.txt']
                print (correo.enviar(asunto = asunto, mensaje = mensaje, receptores = receptores, adjuntos = adjuntos))
"""
"""
    NOTA.- los archivos adjuntos son SU RUTAAAAA no sólo el nombre del archivo salvo que se encuentren en el mismo 
           directorio que el módulo del correo
"""

import os
import smtplib
import string
import email

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import formatdate
from email.encoders import encode_base64

class Mail():

    def __init__(self):
        __version__ = '1.0'
        __VERSION__ = __version__
        __author__ = "Cuahutli Miguel Ulloa Robles"
        self.msgRaiz = MIMEMultipart('related')

    def enviar(self,
    	servidor = 'mail.gruposcit.com',
    	puerto = 26,
    	usuario = 'samplemail@gruposcit.com',
    	password = 'samplepassword',
        receptores=[],
        asunto = 'Mail de prueba desde python script',
        mensaje = 'Este es un email de prueba que se ha enviado desde python',
        adjuntos= []):

        self.servidor = servidor
        self.puerto = puerto
        self.usuario = usuario
        self.password = password
        self.receptores = receptores
        self.asunto = asunto
        self.mensaje = mensaje
        self.adjuntos = adjuntos

        assert type(self.adjuntos) == list

        if self.adjuntos:
            for adjunto in self.adjuntos:
                if not os.path.exists(adjunto):
                    resultado = "error en la ruta o nombre de los archivos adjuntos verifique por favor"
                    return resultado
        x = 0
        recibira = ''
        #genera la cadena de receptores para el correo
        for receptor in self.receptores:
            if not x:
                recibira += receptor 
                x = 1
            else:
                recibira += ','+ receptor

        msgAlternativo = MIMEMultipart('alternative')
        self.msgRaiz.attach(msgAlternativo)

        #Abrimos mensaje html alternativo y lo añadimos
        msgHtml = MIMEText(self.mensaje, 'html', 'utf-8')
        msgAlternativo.attach(msgHtml)

        #Añadimos los ficheros adjuntos a mandar , si los hay
        for archivo in self.adjuntos:
            adjunto = MIMEBase('application', "octet-stream")
            adjunto.set_payload(open(archivo, "rb").read())
            encode_base64(adjunto)
            adjunto.add_header('Content-Disposition', 'attachment; filename = "%s"' %  os.path.basename(archivo))
            self.msgRaiz.attach(adjunto)

        self.msgRaiz['From'] = self.usuario
        self.msgRaiz['To'] = recibira
        self.msgRaiz['Subject'] = self.asunto
        #conexion al servidor de correo
        serverSMTP = smtplib.SMTP(self.servidor,self.puerto)
        serverSMTP.ehlo()
        serverSMTP.starttls()
        serverSMTP.ehlo()
        serverSMTP.login(self.usuario,self.password)

        #enviar Correo
        try:
            serverSMTP.sendmail(self.usuario, recibira.split(','), self.msgRaiz.as_string().encode('ascii'))
            resultado = "Correo enviado satisfactoriamente a:  " + recibira
        except (Exception, e):
            print (Exception, e)
            resultado = "El envío del correo electrónico ha fallado, verifique su información"

        #Cerramos la conexion
        serverSMTP.close()       
        return resultado
        
if __name__== "__main__":
	correo = Mail()
	print (correo.enviar(receptores = ['cm.ulloa@gruposcit.com'], mensaje='Hecho, lo hemos enviado desde python', asunto='Lo hemos enviado con un script en python'))
    	


        
        
       




