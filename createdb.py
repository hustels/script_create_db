# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import unidecode
import xlrd
import mysql.connector
file = 'files/contactos.xls'

workbook = xlrd.open_workbook(file , encoding_override="utf-8") # Abrir el archivo excel
sheet = workbook.sheet_by_index(0)# Eliger la hoja con el que vamos a trabajar
print sheet.cell_value(0,2) # index del valor fila e index del valor columna
#print sheet.nrows #numero de filas
#print sheet.ncols #numero de columnas



		
#Database connection
conn = mysql.connector.connect(user='root', password='root' , host='localhost' , database='contactos')
# Crear el cursor
myCursor = conn.cursor()


for row in range(sheet.nrows):
	atencion = sheet.cell_value(row , 0).encode('ascii', 'ignore').decode('ascii')
	proveedor = sheet.cell_value(row , 1).encode('ascii', 'ignore').decode('ascii')
	contacto = sheet.cell_value(row , 2).encode('ascii', 'ignore').decode('ascii')
	descripcion = sheet.cell_value(row , 3).encode('ascii', 'ignore').decode('ascii')
	#print sheet.cell_value(row , 4).encode('ascii', 'ignore').decode('ascii')
	print unidecode(sheet.cell_value(row , 4).strip())
	














