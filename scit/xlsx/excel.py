#! /usr/bin/ python
# -*- coding: utf-8 -*-

'''
I.S.C. Cuahutli Miguel Ulloa Robles
'''
import collections

import xlsxwriter
from datetime import date, datetime
from decimal import Decimal
from collections import OrderedDict

class Excel:
    
    ERR_NOMNULO     = 'El nombre del archivo del libro no ha sido asignado'
    ERR_CERRADO     = 'El libro ya ha sido cerrado'
    ERR_ESTHOJA     = 'La estructura de la hoja es incorrecto, debe contener los keys llamados "{0}" y "{1}" (hoja: "{2}")'
    ERR_ESTENCA     = 'La definición del encabezado se debe hacer en un list (hoja: "{0}")'
    ERR_ESTDATO     = 'La definición de los datos se debe hacer en un list (hoja: "{0}")'
    ERR_ESTRENG     = 'La definición de cada renglón de los datos se debe hacer en un list (hoja: "{0}")'
    ERR_HOJAS       = 'El Libro que usted intenta analizar no tiene Hojas'
    
    EST_ENCABEZADO  = {'bold': True, 'valign':'vcenter', 'border':1, 'bg_color':'#C6EFCE', 'font_color':'#006100', 'font_size':9}
    EST_GENERAL     = {'border':1, 'font_size':9}
    EST_NUMERO      = {'border':1,  'num_format':'#,##0.00', 'font_size':9}
    EST_SUMATORIAS  = {'border': 0, 'num_format': '#,##0.00','font_color':'#e31a1a', 'bold':True, 'bottom':3, 'font_size':9 }
    EST_FECHA       = {'border':1,  'num_format':'dd/mm/yyyy', 'font_size':9}
    
    def __init__(self, ruta_libro=None):
        if ruta_libro is None:
            raise Exception(Excel.ERR_NOMNULO)
        self._cerrado = False
        self._ruta_libro = ruta_libro
        self._libro = xlsxwriter.Workbook(ruta_libro)
        self._ENCABEZADO = self._libro.add_format(Excel.EST_ENCABEZADO)
        self._GENERAL = self._libro.add_format(Excel.EST_GENERAL)
        self._NUMERO = self._libro.add_format(Excel.EST_NUMERO)
        self._SUMATORIAS = self._libro.add_format(Excel.EST_SUMATORIAS)
        self._FECHA = self._libro.add_format(Excel.EST_FECHA)

    #se agregó header y footer como parámetro para poderle indicar un header y footer por default
    def agregarHoja(self, nombre=None, encabezados=[], datos=[], header=["&C&F - &A"], footer="&CPagina &P de &N" , horizontal = False, tam_hoja= 1,  margen = 0.7, autofiltrar= True ):
        assert isinstance(nombre, str)
        if self._cerrado:
            raise Exception(Excel.ERR_CERRADO)
        nombre = 'Hoja' + str(len(self._libro.worksheets()) + 1) if nombre is None or nombre == '' else nombre
        hoja = self._libro.add_worksheet(nombre)
        hoja.set_margins(margen, margen, 0.75, 0.75)
        if datos:
            c = 0
            r = 0
            x = 1 if encabezados else 0
            aux = []
            for enc in encabezados:
                hoja.write(0, c, enc, self._ENCABEZADO)
                c += 1
            for ren in range(len(datos)):
                aux.append([0] * len(datos[r]))
                c = 0
                if type(datos) == collections.OrderedDict:
                    data = datos[r].values()
                else:
                    data = datos[r]
                for d in data:
                    estilo = None
                    if d == None:
                        estilo = self._GENERAL
                    if type(d) in (str,):
                        estilo = self._GENERAL
                    if type(d) in (datetime, datetime.date, date):
                        estilo = self._FECHA
                    if type(d) in (Decimal, int, float):
                        estilo = self._NUMERO if '.' in str(d) else self._GENERAL
                    hoja.write(r + x, c, d, estilo)
                    ## lista auxiliar para calcular longitud columnas
                    aux[ren][c] = len(str(d)) * 1.1
                    c += 1
                r += 1
            ## magia para las columnas ajustables
            max_w = [list(i) for i in zip(*aux)]
            e_i = 0
            for i in range(len(max_w)):
                if len(encabezados) > i:
                    max_w[i].append(len(encabezados[e_i]) + 5)
                    e_i += 1
                hoja.set_column(i, i, max(max_w[i]))
                #print max_w[i]
        if len(header) >1:
            hoja.set_header(header[0], header[1])
        else:
            hoja.set_header(header[0])
        hoja.set_footer(footer)
        hoja.set_paper(tam_hoja)
        if horizontal:
            hoja.set_landscape()

        hoja.repeat_rows(0) ## repite la primer fila del libro para imprimir
        if autofiltrar:
            hoja.autofilter(0, 0, r - 1, c - 1)

        hoja.freeze_panes(1, 0)

    def modificaMargenes(self, hoja=None, izquierdo=0.7, derecho=0.7, arriba=0.75, abajo=0.75):
        hojas = self.obtenerHojas()
        assert isinstance(hojas, collections.OrderedDict)
        if hoja in hojas:
            sheet = hojas[hoja]
            sheet.set_margins(top=arriba, bottom=abajo, left=izquierdo, right=derecho)


    ## se agregó header y footer como parámetros
    def escribirLibro(self, contenido={}, k_encabezados='encabezados', k_datos='datos', header= ['&C&F - &A'], footer='&CPagina &P de &N', horizontal = False, tam_hoja= 1, margen = 0.7, autofiltrar= True):
        print ("Generando Libro Excel...")
        self.comprobarEstructura(contenido, k_encabezados, k_datos)
        for hoja in contenido.keys():
            self.agregarHoja(hoja, contenido[hoja][k_encabezados], contenido[hoja][k_datos], header, footer, horizontal,tam_hoja, margen, autofiltrar)
    
    def comprobarEstructura(self, contenido, k_encabezados, k_datos):
        for hoja in contenido.keys():
            assert isinstance(hoja, str)
            if k_encabezados not in contenido[hoja].keys() or k_datos not in contenido[hoja].keys():
                raise Exception(Excel.ERR_ESTHOJA.format(k_encabezados, k_datos, hoja))
            if not isinstance(contenido[hoja][k_encabezados], list):
                raise Exception(Excel.ERR_ESTENCA.format(hoja))
            if not isinstance(contenido[hoja][k_datos], list):
                raise Exception(Excel.ERR_ESTDATO.format(hoja))
            else:
                for ren in range(len(contenido[hoja][k_datos])):
                    if not type(contenido[hoja][k_datos][ren]) in (collections.OrderedDict, list):
                        raise Exception(Excel.ERR_ESTRENG.format(hoja))
    
    def obtenerHojas(self):
        if not self._libro.worksheets():
            print (Excel.ERR_HOJAS)
            return Excel.ERR_HOJAS
        hojas= collections.OrderedDict()
        for worksheet in self._libro.worksheets():
            hojas[worksheet.get_name()]=worksheet
        #print (hojas, isinstance(hojas, collections.OrderedDict))
        return hojas

    def modificaHeaderLibro(self, header):
        hojas = self.obtenerHojas()
        assert isinstance(hojas, collections.OrderedDict)
        for hoja in hojas:
            sheet = hojas[hoja]
            sheet.set_header(header)

    def modificaHeader(self, hoja, header):
        hojas = self.obtenerHojas()
        assert isinstance(hojas, collections.OrderedDict)
        if hoja in hojas.keys():
            sheet = hojas[hoja]
            sheet.set_header("&Cheader si, soy header")

    def modificaFooterLibro(self, footer):
        hojas = self.obtenerHojas()
        assert isinstance(hojas, collections.OrderedDict)
        for hoja in hojas:
            sheet = hojas[hoja]
            sheet.set_footer(footer)

    def modificaFooter(self, hoja, footer):
        hojas = self.obtenerHojas()
        assert isinstance(hojas, collections.OrderedDict)
        if hoja in hojas.keys():
            sheet = hojas[hoja]
            sheet.set_header("&CFooter si, soy footer")

    def cerrarLibro(self):
        self._libro.close()
        self._cerrado = True
        
if __name__ == '__main__':
    e = Excel('prueba.xlsx')
    contenido = OrderedDict()
    contenido['1'] = {
        'head': ['enc1.3'],
        'data': [
            ['uno','dos',3],
            ['cuatro',1232132.454,'seis'],
            ['...',date(2014,12,31), date(2000,1,1)]
        ]
    }
    contenido['hoja_prueba2'] = {
        'head': ['enc2.1', 'enc2.2', 'enc2.3'],
        'data': [
            ['uno','dos',3],
            ['cuatro',1232132.454,'seis'],
            ['...',date(2014,12,31), date(2000,1,1)]
        ]
    }
    # contenido = {
    #     '1': {
    #         'head': ['enc1.3'],
    #         'data': [
    #             ['uno','dos',3],
    #             ['cuatro',1232132.454,'seis'],
    #             ['...',date(2014,12,31), date(2000,1,1)]
    #         ]
    #     },
    #     'hoja_prueba2': {
    #         'head': ['enc2.1', 'enc2.2', 'enc2.3'],
    #         'data': [
    #             ['uno','dos',3],
    #             ['cuatro',1232132.454,'seis'],
    #             ['...',date(2014,12,31), date(2000,1,1)]
    #         ]
    #     }
    # }
    try:
        e.escribirLibro(contenido, 'head', 'data')
        # e.agregarHoja('10', ['COL1', 'COL2'], [[1, 2], [3, 4]], "&LFunciona cambiandolo", "yeah vamos!!", True,5)
        # e.obtenerHojas()
    except Exception as ex:
        print ('Ha ocurrido un error durante la creación del libro:\n' + str(ex))
    e.cerrarLibro()
