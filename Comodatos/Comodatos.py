import pandas as pd #'Pandas'
import schedule #'Para automatizar procesos en funcion del dia'
import time 
import glob #'Para buscar el ultimo archivo en una carpeta'
import os
import pyexcel as pe
from datetime import datetime
from tabulate import tabulate
from pandas.io.formats import format as pd_format
import re
import sys
import subprocess
from datetime import datetime, timedelta

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

# CARGAMOS EL ARCHIVO CON SALDOS INICIALES

archivo_excel = pd.ExcelFile('../Comodatos/Cortes/Saldos_Iniciales.xlsx')
saldo_inicial_preventa = archivo_excel.parse('PREVENTA')
saldo_inicial_preventa['cliente'] = saldo_inicial_preventa['cliente'].astype(str)

saldo_inicial_SMK = archivo_excel.parse('SMK')
saldo_inicial_SMK['cliente'] = saldo_inicial_SMK['cliente'].astype(str).str.zfill(5)

saldo_inicial_FLETEROS = archivo_excel.parse('FLETEROS')
saldo_inicial_FLETEROS['cliente'] = saldo_inicial_FLETEROS['cliente'].astype(str).str.zfill(5)

# FECHA DE SALDOS INICIALES

fecha_saldo_inicial_prev = saldo_inicial_preventa.columns[1].strftime('%Y-%m-%d')
fecha_saldo_inicial_SMK = saldo_inicial_SMK.columns[1].strftime('%Y-%m-%d')
fecha_saldo_inicial_FLETEROS = saldo_inicial_SMK.columns[1].strftime('%Y-%m-%d')

directorio_archivos = r'\\layla\\Documentos\\STOCK\\ReporteVacios'

patron_archivo = os.path.join(directorio_archivos, 'Vacios*.xlsx')

archivos = glob.glob(patron_archivo)

archivos_ordenados = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True)

archivo_comodatos = archivos_ordenados[0]
columnas= [ 'articulo', 'cliente', 'fecha', 'haber', 'debe']
df_comodatos = pd.read_excel(archivo_comodatos, usecols = columnas)

# Limpiar caracteres no numéricos de la columna articulos

df_comodatos['articulo'] = df_comodatos['articulo'].apply(lambda x: re.sub('[^0-9]', '', str(x)))

df_comodatos['articulo'] = df_comodatos['articulo'].replace('', '0')
df_comodatos['articulo'] = df_comodatos['articulo'].astype(float).astype(pd.Int64Dtype())

#Arreglamos el formato de la fecha

def cambiar_formato_fecha(fecha_str):
    fecha_obj = datetime.strptime(fecha_str, "%d-%b-%y")
    return fecha_obj.strftime("%d-%m-%Y")

df_comodatos['fecha'] = df_comodatos['fecha'].apply(cambiar_formato_fecha)

articulos_especificos = [9346, 9389]

df_comodatos = df_comodatos[df_comodatos['articulo'].isin(articulos_especificos)]

df_comodatos['fecha'] = pd.to_datetime(df_comodatos['fecha'], format='%d-%m-%Y')
df_filtrado = df_comodatos[df_comodatos['fecha'] >= fecha_saldo_inicial_prev]
movimientos_sum = df_filtrado.groupby('cliente').agg({'debe': 'sum', 'haber': 'sum'}).reset_index()

bd_clientes = pd.read_excel('../Comodatos/bd_clientes.xlsx')

movimientos_sum.sort_values(by = 'debe', ascending= False)

#CLIENTES PREVENTA

cl_preventa = bd_clientes[bd_clientes['CANAL'] == 'PREVENTA']

#CLIENTES MERCADOS

cl_SMK = bd_clientes[bd_clientes['CANAL'] == 'SMK']

#FLETEROS

cl_FLETEROS = bd_clientes[bd_clientes['CANAL'] == 'FLETERO']

#CLIENTES PREVENTA

cl_preventa = cl_preventa[['Codigo', 'DESCRIPCION']]
cl_preventa['Codigo'] = cl_preventa['Codigo'].astype(str)
movimientos_sum['cliente'] = movimientos_sum['cliente'].astype(str)
cl_preventa.rename(columns={'Codigo' : 'cliente'}, inplace=True) 

acum_preventa = cl_preventa.merge(movimientos_sum, on= 'cliente', how = 'left')
acum_preventa.fillna(0, inplace= True)

# UNIMOS SALDOS INICIALES CON EL DF DE LOS MOVIMIENTOS

preventa_actualizado = acum_preventa.merge(saldo_inicial_preventa, on = 'cliente', how= 'left')

# REEMPLAZAMOS LOS NULOS POR 0
preventa_actualizado.fillna(0, inplace = True)

# SACAMOS LA FECHA DEL SALDO INICIAL

fecha_saldo_inicial =  preventa_actualizado.columns[4]

fecha_saldo_formateada = preventa_actualizado.columns[4].strftime('%Y-%m-%d')

# ORDENAMOS LAS COLUMNAS A NECESIDAD

orden_columnas = ['cliente', 'DESCRIPCION', fecha_saldo_inicial, 'debe', 'haber']
preventa_actualizado = preventa_actualizado.reindex(columns= orden_columnas)

# AGREGAMOS LA COLUMNA SALDO ACTUALIZADO, QUE SERA IGUAL AL SALDO INCIIAL + DEBE - HABER
fecha_acutal = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
nombre_saldo_actualizado = 'Saldo actualizado al ' + fecha_acutal 
preventa_actualizado[nombre_saldo_actualizado] = preventa_actualizado[fecha_saldo_inicial] + preventa_actualizado['debe'] - preventa_actualizado['haber']
preventa_actualizado=preventa_actualizado.rename(columns= {fecha_saldo_inicial : fecha_saldo_formateada})
preventa_actualizado=preventa_actualizado.sort_values(by= nombre_saldo_actualizado, ascending = False)

directorio_archivos = r'H:\\STOCK\\COMODATOS VACIOS\\PREVENTA\\'

patron_archivo = os.path.join(directorio_archivos, 'Comodatos-Preventa-*.xlsx')

archivos = glob.glob(patron_archivo)

archivos_ordenados = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True)

archivo_cmd = pd.ExcelFile(archivos_ordenados[0])

df_evolucion = archivo_cmd.parse('Evolucion')

df_evolucion = df_evolucion.iloc[:, :-1]

# Obtener la lista de todas las columnas de saldo

columnas_saldo = df_evolucion.columns[2:]  # Suponiendo que las dos primeras columnas son 'cliente' y 'descrip'

# Obtener la diferencia entre el saldo actualizado y los saldos

preventa_actualizado = preventa_actualizado.reset_index(drop= True)

df_evolucion[fecha_acutal] = preventa_actualizado[nombre_saldo_actualizado] - df_evolucion[columnas_saldo].sum(axis=1)

columnas_saldo = df_evolucion.columns[2:]

df_evolucion['Saldo_actualizado'] = df_evolucion[columnas_saldo].sum(axis=1)

# GENERAR EL ARCHIVO CON LOS DATOS ACTUALIZADOS

nombre_archivo = 'Comodatos-Preventa-Actualizados-' + fecha_acutal + '.xlsx'

ruta_carpeta = '\\\\layla\\Documentos\\STOCK\\COMODATOS VACIOS\\PREVENTA\\'

ruta_completa = os.path.join(ruta_carpeta, nombre_archivo)

preventa_actualizado.to_excel(ruta_completa, index=False)

# Crear un objeto ExcelWriter
writer = pd.ExcelWriter(ruta_completa, engine='openpyxl')

# Guardar el DataFrame principal en la hoja 1
preventa_actualizado.to_excel(writer, sheet_name='Preventa', index=False)

# Guardar la tabla resumen en una nueva hoja
df_evolucion.to_excel(writer, sheet_name='Evolucion', index=False)

# Cerrar el escritor
writer.close()

df_comodatos['fecha'] = pd.to_datetime(df_comodatos['fecha'], format='%d-%m-%Y')
df_filtrado_smk = df_comodatos[df_comodatos['fecha'] >= fecha_saldo_inicial_SMK]
movimientos_sum_smk = df_filtrado_smk.groupby('cliente').agg({'debe': 'sum', 'haber': 'sum'}).reset_index()

cl_SMK = cl_SMK[['Codigo', 'DESCRIPCION']] # SOLO NOS INTERESA EL CODIGO Y LA DESCRIPCION
cl_SMK['Codigo'] = cl_SMK['Codigo'].astype(str).str.zfill(5) # NORMALIZAMOS EL CODIGO COMO STR PARA QUE DESPUES FUNCIONE EL JOIN
movimientos_sum_smk['cliente'] = movimientos_sum_smk['cliente'].astype(str) # IDEM ANTERIOR PARA QUE FUNCIONE EL JOIN
cl_SMK.rename(columns={'Codigo' : 'cliente'}, inplace=True) # CAMBIAMOS EL NOMBRE PARA QUE TENGA EL MISMO EN TODAS LAS TABLAS

acum_smk = cl_SMK.merge(movimientos_sum_smk, on= 'cliente', how = 'left') # UNIMOS PARA TRAER LA CANTIDAD A CADA CLIENTE
acum_smk.fillna(0, inplace= True)  #PONEMOS EN 0 LOS NAN
acum_smk['cliente'] = acum_smk['cliente'].astype(str)

# UNIMOS SALDOS INICIALES CON EL DF DE LOS MOVIMIENTOS

smk_actualizado = acum_smk.merge(saldo_inicial_SMK, on = 'cliente', how= 'left')

# REEMPLAZAMOS LOS NULOS POR 0

smk_actualizado.fillna(0, inplace = True)

# SACAMOS LA FECHA DEL SALDO INICIAL

fecha_saldo_inicial =  smk_actualizado.columns[4]

fecha_saldo_formateada = smk_actualizado.columns[4].strftime('%Y-%m-%d')

# ORDENAMOS LAS COLUMNAS A NECESIDAD

orden_columnas = ['cliente', 'DESCRIPCION', fecha_saldo_inicial, 'debe', 'haber']
smk_actualizado = smk_actualizado.reindex(columns= orden_columnas)

# AGREGAMOS LA COLUMNA SALDO ACTUALIZADO, QUE SERA IGUAL AL SALDO INCIIAL + DEBE - HABER
fecha_acutal = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
nombre_saldo_actualizado = 'Saldo actualizado al ' + fecha_acutal 
smk_actualizado[nombre_saldo_actualizado] = smk_actualizado[fecha_saldo_inicial] + smk_actualizado['debe'] - smk_actualizado['haber']
smk_actualizado = smk_actualizado.rename(columns= {fecha_saldo_inicial : fecha_saldo_formateada})
smk_actualizado = smk_actualizado.sort_values(by= nombre_saldo_actualizado, ascending = False)

# Calcular el saldo total por mercado
tabla_resumen = smk_actualizado.groupby('DESCRIPCION')[nombre_saldo_actualizado].sum().reset_index().sort_values(by = nombre_saldo_actualizado, ascending= False)

print(tabla_resumen)

directorio_archivos = r'H:\\STOCK\\COMODATOS VACIOS\\MERCADO\\'

patron_archivo = os.path.join(directorio_archivos, 'Comodatos-SMK-*.xlsx')

archivos = glob.glob(patron_archivo)

archivos_ordenados = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True)

archivo_cmd = pd.ExcelFile(archivos_ordenados[0])

df_evolucion = archivo_cmd.parse('EvolucionSMK')

df_evolucion = df_evolucion.iloc[:, :-1]

# Obtener la lista de todas las columnas de saldo

columnas_saldo = df_evolucion.columns[2:]  # Suponiendo que las dos primeras columnas son 'cliente' y 'descrip'

# Obtener la diferencia entre el saldo actualizado y los saldos

smk_actualizado = smk_actualizado.reset_index(drop= True)

df_evolucion[fecha_acutal] = smk_actualizado[nombre_saldo_actualizado] - df_evolucion[columnas_saldo].sum(axis=1)

columnas_saldo = df_evolucion.columns[2:]

df_evolucion['Saldo_actualizado'] = df_evolucion[columnas_saldo].sum(axis=1)

from openpyxl import Workbook

# GENERAR EL ARCHIVO CON LOS DATOS ACTUALIZADOS

nombre_archivo = 'Comodatos-SMK-Actualizados-' + fecha_acutal + '.xlsx'

ruta_carpeta = '\\\\layla\\Documentos\\STOCK\\COMODATOS VACIOS\\MERCADO\\'

ruta_completa = os.path.join(ruta_carpeta, nombre_archivo)

# Crear un objeto ExcelWriter
writer = pd.ExcelWriter(ruta_completa, engine='openpyxl')

# Guardar el DataFrame principal en la hoja 1
smk_actualizado.to_excel(writer, sheet_name='Detalle SMK', index=False)

# Guardar la tabla resumen en una nueva hoja
tabla_resumen.to_excel(writer, sheet_name='Tabla Resumen', index=False)

# Guardar la tabla resumen en una nueva hoja
df_evolucion.to_excel(writer, sheet_name='EvolucionSMK', index=False)

# Cerrar el escritor
writer.close()

df_comodatos['fecha'] = pd.to_datetime(df_comodatos['fecha'], format='%d-%m-%Y')
df_filtrado_fleteros = df_comodatos[df_comodatos['fecha'] >= fecha_saldo_inicial_FLETEROS]
movimientos_sum_FLETEROS = df_filtrado_fleteros.groupby('cliente').agg({'debe': 'sum', 'haber': 'sum'}).reset_index()

cl_FLETEROS = cl_FLETEROS[['Codigo', 'DESCRIPCION']] # SOLO NOS INTERESA EL CODIGO Y LA DESCRIPCION
cl_FLETEROS['Codigo'] = cl_FLETEROS['Codigo'].astype(str).str.zfill(5) # NORMALIZAMOS EL CODIGO COMO STR PARA QUE DESPUES FUNCIONE EL JOIN
movimientos_sum_FLETEROS['cliente'] = movimientos_sum_FLETEROS['cliente'].astype(str) # IDEM ANTERIOR PARA QUE FUNCIONE EL JOIN
cl_FLETEROS.rename(columns={'Codigo' : 'cliente'}, inplace=True) # CAMBIAMOS EL NOMBRE PARA QUE TENGA EL MISMO EN TODAS LAS TABLAS

acum_FLETEROS = cl_FLETEROS.merge(movimientos_sum_FLETEROS, on= 'cliente', how = 'left') # UNIMOS PARA TRAER LA CANTIDAD A CADA CLIENTE
acum_FLETEROS.fillna(0, inplace= True)  #PONEMOS EN 0 LOS NAN
acum_FLETEROS['cliente'] = acum_FLETEROS['cliente'].astype(str)

# UNIMOS SALDOS INICIALES CON EL DF DE LOS MOVIMIENTOS

FLETEROS_actualizado = acum_FLETEROS.merge(saldo_inicial_FLETEROS, on = 'cliente', how= 'left')

# REEMPLAZAMOS LOS NULOS POR 0

FLETEROS_actualizado.fillna(0, inplace = True)

# SACAMOS LA FECHA DEL SALDO INICIAL

fecha_saldo_inicial =  FLETEROS_actualizado.columns[4]

fecha_saldo_formateada = FLETEROS_actualizado.columns[4].strftime('%Y-%m-%d')

# ORDENAMOS LAS COLUMNAS A NECESIDAD

orden_columnas = ['cliente', 'DESCRIPCION', fecha_saldo_inicial, 'debe', 'haber']
FLETEROS_actualizado = FLETEROS_actualizado.reindex(columns= orden_columnas)

# AGREGAMOS LA COLUMNA SALDO ACTUALIZADO, QUE SERA IGUAL AL SALDO INCIIAL + DEBE - HABER

fecha_acutal = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
nombre_saldo_actualizado = 'Saldo actualizado al ' + fecha_acutal 
FLETEROS_actualizado[nombre_saldo_actualizado] = FLETEROS_actualizado[fecha_saldo_inicial] + FLETEROS_actualizado['debe'] - FLETEROS_actualizado['haber']
FLETEROS_actualizado = FLETEROS_actualizado.rename(columns= {fecha_saldo_inicial : fecha_saldo_formateada})
FLETEROS_actualizado = FLETEROS_actualizado.sort_values(by= nombre_saldo_actualizado, ascending = False)

directorio_archivos = r'H:\\STOCK\\COMODATOS VACIOS\\FLETEROS\\'

patron_archivo = os.path.join(directorio_archivos, 'Comodatos-Fleteros*.xlsx')

archivos = glob.glob(patron_archivo)

archivos_ordenados = sorted(archivos, key=lambda x: os.path.getmtime(x), reverse=True)

archivo_cmd = pd.ExcelFile(archivos_ordenados[0])

df_evolucion = archivo_cmd.parse('Evolucion')

df_evolucion = df_evolucion.iloc[:, :-1]

# Obtener la lista de todas las columnas de saldo

columnas_saldo = df_evolucion.columns[2:]  # Suponiendo que las dos primeras columnas son 'cliente' y 'descrip'

# Obtener la diferencia entre el saldo actualizado y los saldos

FLETEROS_actualizado = FLETEROS_actualizado.reset_index(drop= True)

df_evolucion[fecha_acutal] = FLETEROS_actualizado[nombre_saldo_actualizado] - df_evolucion[columnas_saldo].sum(axis=1)

columnas_saldo = df_evolucion.columns[2:]

df_evolucion['Saldo_actualizado'] = df_evolucion[columnas_saldo].sum(axis=1)

# GENERAR EL ARCHIVO CON LOS DATOS ACTUALIZADOS

from openpyxl import Workbook

nombre_archivo = 'Comodatos-Fleteros-Actualizados-' + fecha_acutal + '.xlsx'

ruta_carpeta = '\\\\layla\\Documentos\\STOCK\\COMODATOS VACIOS\\FLETEROS\\'

ruta_completa = os.path.join(ruta_carpeta, nombre_archivo)

FLETEROS_actualizado.to_excel(ruta_completa, index=False)

# Crear un objeto ExcelWriter
writer = pd.ExcelWriter(ruta_completa, engine='openpyxl')

# Guardar el DataFrame principal en la hoja 1
FLETEROS_actualizado.to_excel(writer, sheet_name='Fleteros', index=False)

# Guardar la tabla resumen en una nueva hoja
df_evolucion.to_excel(writer, sheet_name='Evolucion', index=False)

# Cerrar el escritor
writer.close()