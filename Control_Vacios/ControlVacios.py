import pandas as pd #'Pandas'
import time 
import glob #'Para buscar el ultimo archivo en una carpeta'
import os
import pyexcel as pe
from datetime import datetime
from tabulate import tabulate
from pandas.io.formats import format as pd_format
import os
import sys
import subprocess

### VENTAS ###

directorio_archivos = r"\\layla\Documentos\STOCK\ReporteVacios"

patron_archivo = os.path.join(directorio_archivos, 'Ventas*.xls') #Aca va el tipo de archivo

archivos = glob.glob(patron_archivo)

archivos = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True) #Ordenar archivos por el ultimo modificado

archivo_excel_original = archivos[0]

archivo_excel_destino = archivo_excel_original + "x"

# Comando para abrir el archivo de Excel original
comando_abrir = 'start "{}"'.format(archivo_excel_original)

# Ejecutar el comando para abrir el archivo de Excel original
subprocess.run(comando_abrir, shell=True)

# Esperar unos segundos para asegurarse de que Excel esté abierto y listo
# Ajusta este tiempo según sea necesario
time.sleep(2)

# Comando de PowerShell para guardar el archivo de Excel como .xlsx
comando_guardar = 'powershell -c "$excel = New-Object -ComObject Excel.Application; $workbook = $excel.Workbooks.Open(\'{}\'); $workbook.SaveAs(\'{}\', 51); $excel.Quit()"'.format(archivo_excel_original, archivo_excel_destino)

# Ejecutar el comando para guardar el archivo de Excel como .xlsx
subprocess.run(comando_guardar, shell=False)

# DESCARGAS

directorio_archivos = r"\\layla\Documentos\STOCK\ReporteVacios"

patron_archivo = os.path.join(directorio_archivos, 'Descargas*.xls') #Aca va el tipo de archivo

archivos = glob.glob(patron_archivo)

archivos = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True) #Ordenar archivos por el ultimo modificado

archivo_excel_original = archivos[0]

archivo_excel_destino = archivo_excel_original + "x"

# Comando para abrir el archivo de Excel original
comando_abrir = 'start "{}"'.format(archivo_excel_original)

# Ejecutar el comando para abrir el archivo de Excel original
subprocess.run(comando_abrir, shell=True)

# Esperar unos segundos para asegurarse de que Excel esté abierto y listo
# Ajusta este tiempo según sea necesario
time.sleep(2)

# Comando de PowerShell para guardar el archivo de Excel como .xlsx
comando_guardar = 'powershell -c "$excel = New-Object -ComObject Excel.Application; $workbook = $excel.Workbooks.Open(\'{}\'); $workbook.SaveAs(\'{}\', 51); $excel.Quit()"'.format(archivo_excel_original, archivo_excel_destino)

# Ejecutar el comando para guardar el archivo de Excel como .xlsx
subprocess.run(comando_guardar, shell=False)


# COMODATOS

directorio_archivos = r"\\layla\Documentos\STOCK\ReporteVacios"

patron_archivo = os.path.join(directorio_archivos, 'Vacios*.xls') #Aca va el tipo de archivo

archivos = glob.glob(patron_archivo)

archivos = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True) #Ordenar archivos por el ultimo modificado

archivo_excel_original = archivos[0]

archivo_excel_destino = archivo_excel_original + "x"

# Comando para abrir el archivo de Excel original
comando_abrir = 'start "{}"'.format(archivo_excel_original)

# Ejecutar el comando para abrir el archivo de Excel original
subprocess.run(comando_abrir, shell=True)

# Esperar unos segundos para asegurarse de que Excel esté abierto y listo
# Ajusta este tiempo según sea necesario
time.sleep(2)

# Comando de PowerShell para guardar el archivo de Excel como .xlsx
comando_guardar = 'powershell -c "$excel = New-Object -ComObject Excel.Application; $workbook = $excel.Workbooks.Open(\'{}\'); $workbook.SaveAs(\'{}\', 51); $excel.Quit()"'.format(archivo_excel_original, archivo_excel_destino)

# Ejecutar el comando para guardar el archivo de Excel como .xlsx
subprocess.run(comando_guardar, shell=False)

patron_archivo = 'VentasDesde*.xlsx'

directorio_archivos = r'\\layla\\Documentos\\STOCK\\ReporteVacios'

patron_archivo = os.path.join(directorio_archivos, 'Ventas*.xlsx')

archivos = glob.glob(patron_archivo)

archivos_ordenados = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True)

archivo_ventas = archivos_ordenados[0]
columnas= ['fecha', 'reparto', 'entrada', 'salida', 'articulo_a', 'cliente']
df_ventas = pd.read_excel(archivo_ventas, usecols = columnas)
df_vent_clientes = df_ventas

articulos_especificos = [9951, 19099, 20789, 20791, 21149, 18873, 23926, 20993, 9376, 21836, 9838, 9364, 9953
, 22475, 22643, 20826]

df_ventas = df_ventas[df_ventas['articulo_a'].isin(articulos_especificos)]

# Creamos la funcion para transformar la fecha

def cambiar_formato_fecha(fecha_str):
    fecha_obj = datetime.strptime(fecha_str, "%d-%b-%y")
    return fecha_obj.strftime("%d-%m-%Y")

# Aplica la función a cada valor de la columna "fecha"

df_ventas['fecha'] = df_ventas['fecha'].apply(cambiar_formato_fecha)
df_vent_clientes['fecha'] = df_vent_clientes['fecha'].apply(cambiar_formato_fecha)

df_ventas['saldo'] = df_ventas['salida']

# De la fecha tomamos el valor minimo que es el dia en el que el reparto fue abierto, y el saldo lo sumamos para acumularlo

df_ventas = df_ventas.groupby(['reparto']).agg({'saldo': 'sum', 'fecha' : 'min'}).reset_index() # Agrupar por reparto, agregando la suma del saldo y el valor minimo de fecha
df_ventas = df_ventas.sort_values('fecha', ascending = False) # Ordenar por fecha
df_ventas['fecha'] = pd.to_datetime(df_ventas['fecha']).dt.date # Pasar la fecha al formato correcto
columnas = ['fecha', 'reparto', 'saldo'] # Aca ordenamos las columnas en el orden deseado
df_ventas = df_ventas.reindex(columns=columnas) # Idem aca para odenar

from datetime import datetime, timedelta
fecha_actual = datetime.now().date()
fecha_limite = fecha_actual - timedelta(days=7)
df_ventas = df_ventas[(df_ventas['fecha'] >= fecha_limite) & (df_ventas['fecha'] <= fecha_actual)]

directorio_archivos = r'\\layla\\Documentos\\STOCK\\ReporteVacios'

patron_archivo = os.path.join(directorio_archivos, 'Descargas*.xlsx')

archivos = glob.glob(patron_archivo)

archivos_ordenados = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True)

archivo_descargas = archivos_ordenados[0]
columnas= ['reparto', 'entrada', 'salida', 'articulo_a']
df_descargas = pd.read_excel(archivo_descargas, usecols = columnas)

articulos_especificos = [9346, 9376, 9364, 9389, 9838, 9951, 9953, 18873, 19099, 20789, 20791,20826
, 20993, 21149, 21836, 22475, 22642, 22643, 23134, 23926]

df_descargas = df_descargas[df_descargas['articulo_a'].isin(articulos_especificos)]

df_descargas['regreso'] = df_descargas['entrada'] - df_descargas['salida']

# Vamos a agrupar los valores por reparto

# De la fecha tomamos el valor minimo que es el dia en el que el reparto fue abierto, y el saldo lo sumamos para acumularlo

df_descargas = df_descargas.groupby(['reparto']).agg({'regreso': 'sum'}).reset_index() # Agrupar por reparto, agregando la suma del saldo
df_descargas = df_descargas.sort_values('reparto', ascending = False) # Ordenar por reparto
columnas = ['reparto', 'regreso'] # Aca ordenamos las columnas en el orden deseado
df_descargas = df_descargas.reindex(columns=columnas) # Idem aca para odenar

df_merged = df_ventas.merge(df_descargas, on='reparto', how='left') #Unimos con un left join asi nos trae todos los de ventas
df_merged['regreso'] = df_merged['regreso'].fillna(0).astype(int) #Si hay alguno en ventas que no coincide con descargas le pone 0
df_merged = df_merged[df_merged['saldo'] != df_merged['regreso']] #Me quedo solo con los que son distintos, los otros estan OK

#Obtenemos los numeros de los repartos que no coinciden sus saldos

numeros_reparto = df_merged['reparto'].unique().tolist()

directorio_archivos = r'\\layla\\Documentos\\STOCK\\ReporteVacios'

patron_archivo = os.path.join(directorio_archivos, 'Vacios*.xlsx')

archivos = glob.glob(patron_archivo)

archivos_ordenados = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True)

import re

archivo_comodatos = archivos_ordenados[0]
columnas= ['reparto', 'debe', 'articulo', 'cliente', 'fecha']
df_comodatos = pd.read_excel(archivo_comodatos, usecols = columnas)

# Limpiar caracteres no numéricos de la columna articulos
df_comodatos['articulo'] = df_comodatos['articulo'].apply(lambda x: re.sub('[^0-9]', '', str(x)))

df_comodatos['articulo'] = df_comodatos['articulo'].replace('', '0')
df_comodatos['articulo'] = df_comodatos['articulo'].astype(float).astype(pd.Int64Dtype())

# Limpiar caracteres no numéricos de la columna reparto
df_comodatos['reparto'] = df_comodatos['reparto'].apply(lambda x: re.sub('[^0-9]', '', str(x)))

#Nos quedamos con los que tienen reparto vacios (Son los preventa los otros son mercados)

df_comodatos=df_comodatos.dropna(subset=['reparto'])

#Arreglamos el formato de la fecha

df_comodatos['fecha'] = df_comodatos['fecha'].apply(cambiar_formato_fecha)

articulos_especificos = [9346, 9389]

df_comodatos = df_comodatos[df_comodatos['articulo'].isin(articulos_especificos)]

df_comodatos.rename(columns={'debe' : 'comodato'}, inplace=True) #renombramos para que se llame igual a las demas
df_comodatos['fecha'] = pd.to_datetime(df_comodatos['fecha'], format='%d-%m-%Y')

for numero_reparto in numeros_reparto:
    # Filtrar la tabla Ventas por el número de reparto actual
    df_vent = df_vent_clientes[df_vent_clientes['reparto'] == numero_reparto]
    
    # Obtener la lista de clientes únicos en la tabla Ventas del reparto actual
    clientes_reparto = df_vent['cliente'].unique().tolist()
    
    # Filtrar la tabla Comodatos por las fechas dentro del rango de 3 días posteriores a la fecha del reparto actual
    fechas_reparto = df_merged[df_merged['reparto'] == numero_reparto]['fecha']
    fechas_rango = [fecha + pd.DateOffset(days=dias) for fecha in fechas_reparto for dias in range(1, 5)]
    comodatos_reparto = df_comodatos[df_comodatos['fecha'].isin(fechas_rango)]
    
# Calcular la cantidad de comodatos realizados a cada cliente en el reparto actual
    comodatos_cliente = comodatos_reparto[comodatos_reparto['cliente'].isin(clientes_reparto)]
    comodatos_por_cliente = comodatos_cliente.groupby('cliente')['comodato'].sum().reset_index()

    # Verificar si hay al menos una fila en comodatos_por_cliente
    if not comodatos_por_cliente.empty:
        valor_comodato = comodatos_por_cliente['comodato'].values[0]

    # Agregar la columna de comodatos a la tabla Merged
        df_merged.loc[df_merged['reparto'] == numero_reparto, 'comodato'] = valor_comodato
    else:
        df_merged.loc[df_merged['reparto'] == numero_reparto, 'comodato'] = 0  # O cualquier otro valor predeterminado en caso de no haber comodatos

from tabulate import tabulate

# Crear una lista de listas con los datos de la tabla
tabla = []
for _, row in df_merged.iterrows():
    fila = [
        row['fecha'],
        row['reparto'],
        row['saldo'],
        row['regreso'],
        row['comodato'],
        '',
    ]
    tabla.append(fila)

# Configurar opciones de formato
opciones = {
    'headers': ['Fecha', 'Reparto', 'Salida', 'Regreso', 'Comodato', 'Verificación'],
    'tablefmt': 'pretty',  # Formato de la tabla (opciones: plain, simple, grid, fancy_grid, ...)
}

import pandas as pd
from fpdf import FPDF

# Crear una clase PDF personalizada
class PDF(FPDF):
    def header(self):
        # Configurar la cabecera del PDF
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Tabla de Datos', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        # Configurar el pie de página del PDF
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, 'Página %s' % self.page_no(), 0, 0, 'C')

    def chapter_title(self, title):
        # Configurar el título del capítulo
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(5)

    def chapter_body(self, data):
        # Configurar el cuerpo del capítulo con los datos de la tabla
        self.set_font('Arial', '', 10)
        self.cell(25, 10, 'Fecha', 1, 0, 'C')
        self.cell(25, 10, 'Reparto', 1, 0, 'C')
        self.cell(25, 10, 'Salida', 1, 0, 'C')
        self.cell(25, 10, 'Regreso', 1, 0, 'C')
        self.cell(25, 10, 'Comodato', 1, 0, 'C')
        self.cell(25, 10, 'Verificación', 1, 1, 'C')

        for index, row in data.iterrows():
            self.cell(25, 10, str(row['fecha']), 1, 0, 'C')
            self.cell(25, 10, str(row['reparto']), 1, 0, 'C')
            self.cell(25, 10, str(row['saldo']), 1, 0, 'C')
            self.cell(25, 10, str(row['regreso']), 1, 0, 'C')
            self.cell(25, 10, str(row['comodato']), 1, 0, 'C')
            self.cell(25, 10, '', 1, 1, 'C')  # Casilla de verificación vacía

# Crear el objeto PDF
pdf = PDF()
pdf.add_page()

# Agregar el contenido al PDF
pdf.chapter_body(df_merged)

# Guardar el PDF en un archivo
pdf_filename = 'tabla_datos.pdf'
pdf.output(pdf_filename)


import subprocess

def imprimir_pdf(pdf_filename):
    try:
        subprocess.Popen(['C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe', '/p', '/h', pdf_filename])
        print("El archivo PDF se ha enviado para imprimir.")
    except Exception as e:
        print("Error al imprimir el archivo PDF:", str(e))

# Llamar a la función de impresión del archivo PDF
pdf_filename = r'C:\Users\mdasilva\Desktop\Contro Vacios\tabla_datos.pdf'
imprimir_pdf(pdf_filename)
