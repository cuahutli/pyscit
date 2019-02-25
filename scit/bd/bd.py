#! /usr/bin/ python
# -*- coding: utf-8 -*-

'''
Created by: Cuahutli Miguel Ulloa Robles
'''

db = None
class Bd:

	ERR_ARRVACI = 'El arreglo de columnas/valores está vacío'
	ERR_ARRLONG = 'La longitud de los nombres de columna es diferente a la de los valores'
	ERR_CONXINI = 'La conexión no ha sido inicializada'
	ERR_ARRVALS = 'El arreglo de valores no puede ser nulo'

	def __init__(self, conexion=None, dbm = ''):
		self._conexion = conexion
		self._dbm = dbm
		self._columnas = []
		
	@classmethod
	def iniParametros(cls, servidor, base, usuario, clave, dbm = 'pgsql'):
		global db
		print(servidor,base,usuario,clave,dbm)
		if dbm not in ('pgsql','mysql'):
			return ("el manejador debe ser pgsql (postgres) o mysql (mysql)")
		if dbm == 'pgsql':
			try:
				import psycopg2 as db
			except ImportError:
				raise ImportError ("no tiene instalado psycopg2, instalelo por favor \n pip install psycopg2")
		else:
			try:
				import mysqldb as db
			except ImportError:
				raise ImportError ("no tiene instalado mysqldb, instalelo por favor \n pip install myslqdb")

		try:
			
			conexion = db.connect(host =servidor,  database= base, user = usuario, password = clave) if dbm == 'pgsql' else db.connect(host =servidor,  db= base, user = usuario, passwd = clave)	
			return cls(conexion,dbm) 
		except db.Error as ex:
			return "Error: Manejador BD [%d]: %s" % (ex.args[0], ex.args[1])

	@classmethod
	def iniConexion(cls, conexion):
		return cls(conexion)


	def cerrarConexion(self):
		self._conexion.close()

	def selTabla(self, tabla):
		sql = '''select * from %s;''' % (tabla)
		return self.select(sql)

	def getEncoding(self):
		return self._conexion.encoding

	def select (self, consulta, *args):
		if self._conexion is None:
			raise Exception(Query.ERR_CONXINI)
		if self._dbm == 'pgsql':
			import psycopg2.extras

		cursor = self._conexion.cursor(cursor_factory=psycopg2.extras.DictCursor) if self._dbm == 'pgsql' else self._conexion.cursor(db.cursors.DictCursor)
		try:
			cursor.execute(consulta,args)
			self._conexion.commit()
			resultado = cursor.fetchall()
			self._columnas = [desc[0] for desc in cursor.description]
			return resultado
		except db.Error as ex:
				self._conexion.rollback()
				return "Error: Manejador BD [%d]: %s" % (ex.args[0], ex.args[1])

	def getColumnas(self):
		return self._columnas

""" 
if __name__ == '__main__':
	from datetime import date
	import json
	m = Bd.iniParametros('localhost', 'database_name', 'db_user', 'user_password')
	r = m.select('select * from tabla where cond > 1')
	print (m.getColumnas())
	print (m.getEncoding())
	print (r)

 """