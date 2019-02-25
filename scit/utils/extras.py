#! /usr/bin/ python
# -*- coding: utf-8 -*-

import os
import locale
import platform

import calendar
from  datetime import datetime, timedelta


def validaDir(dirdestino):
    if not os.path.exists(dirdestino):
        os.makedirs(dirdestino)


def localeEs():
    try:
        locale.setlocale(locale.LC_ALL, 'es_MX.UTF-8')
    except:
        locale.setlocale(locale.LC_ALL, '')  # locale por default en el sistema


def copiaArch(origen, destinos):
    from subprocess import call

    logfd = open('errores_copiar', 'w+')
    for destino in destinos:
        validaDir(destino)
        comando = 'copy "{0}" "{1}" '.format(origen,
                                             destino) if platform.system() == 'Windows' else "cp '{0}' '{1}' ".format(
            origen, destino)
        # print comando
        proceso = call(comando, shell=True, stdout=logfd, stderr=logfd)
        print "Archivo copiado: \n\t de: {0} \n\t a:{1}".format(origen, destino)


''' calcula mes anterior '''


def mesAnterior(hoy):
    if hoy.month == 1:
        mes = 12
        ano = hoy.year - 1
    else:
        mes = hoy.month - 1
        ano = hoy.year
    return ano, mes


''' calcula el primer día del mes actual '''


def primerDia(hoy):
    fec_ini = datetime(hoy.year, hoy.month, 1)
    return fec_ini


''' calcula el primer día del mes anterior '''


def diaMesAnt(hoy):
    fec_ini = datetime(hoy.year - 1, 12, 1) if hoy.month == 1 else datetime(hoy.year, hoy.month - 1, 1)
    return fec_ini


''' calcula el primer y ultima día del mes anterior a la fecha dada '''


def calculaDias(hoy):
    fec_ini = datetime(hoy.year - 1, 12, 1) if hoy.month == 1 else datetime(hoy.year, hoy.month - 1, 1)
    fec_fin = datetime(hoy.year - 1, 12, calendar.monthrange(hoy.year - 1, 12)[1]) if hoy.month == 1 else  datetime(
        hoy.year, hoy.month - 1, calendar.monthrange(hoy.year, hoy.month - 1)[1])
    return fec_ini, fec_fin


def getFirma():
    #path = os.path.expanduser('~') + os.sep + 'scripts' + os.sep + 'assets' + os.sep
    path = 'assets' + os.sep
    #print path
    f = open(path + 'firma', 'r')
    mensaje = '''<p style="font-size:10px;"><i>Este mensaje ha sido enviado automáticamente, no necesita responderlo.</p></i>{0}'''.format(
    f.read())
    return mensaje

if __name__ == "__main__":
    localeEs()
    fecha = datetime(2015, 1, 19)
    fec_ini, fec_fin = calculaDias(fecha)
    print fec_ini
    print fec_fin
