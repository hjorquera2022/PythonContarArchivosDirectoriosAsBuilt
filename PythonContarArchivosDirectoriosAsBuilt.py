#PythonContarArchivosDirectoriosAsBuilt.py
# CuadraturaProvisoria

import pandas as pd
import os
import openpyxl

#
#Estructura de las Carpetas en REVISORES  por cada parcialidad.
#
#├───ESTRUCTURA DE CARPETAS DE CADA PARCIALIDAD
#
#│   ├───DOCUMENTOS ASBUILT 
#       ├───01 PDF
#       │   ├───APROBADOS
#       │   └───OBSERVADOS
#       └───02 EDITABLE
#           ├───APROBADOS
#           └───OBSERVADOS
#                    
# Lista de carpetas y subcarpetas

estructura = [
            'DOCUMENTOS ASBUILT/01 PDF/APROBADOS',
            'DOCUMENTOS ASBUILT/01 PDF/OBSERVADOS',
            'DOCUMENTOS ASBUILT/02 EDITABLE/APROBADOS',
            'DOCUMENTOS ASBUILT/02 EDITABLE/OBSERVADOS',
             ]

# Función para contar archivos en una carpeta y sus subcarpetas
def contar_archivos(ruta):
    contador = 0
    for raiz, directorios, archivos in os.walk(ruta):
        for archivo in archivos:
            contador += 1
    return contador



# Ruta base donde se deben verificar los subdirectorios
ruta_base = 'R:\\01 PARCIALIDADES\\'   

# Nombre del archivo de log
archivo_log = ruta_base + '0000-00 ADMINISTRACION\\LOG\\log_CuadraturaProvisoriaAsBuilt.txt'

# Planilla con la lista de parcialidades
archivo_excel = ruta_base + 'Listado de Parcialidades_AsBuilt.xlsx'

# Carga el archivo Excel en un DataFrame Hoja de Parcialidades.
df = pd.read_excel(archivo_excel, sheet_name='PARCIALIDADES')

# Filtra el DataFrame para considerar solo parcialidades a 'PROCESAR' igual a 'S'
df_parcialidades = df[df['PROCESAR'] == 'S']

# Abre el archivo de log en modo de escritura
with open(archivo_log, 'w') as log_file:

    # Itera a través de cada parcialidad y la procesa
    for parcialidad in df_parcialidades['PARCIALIDAD']:

        #******* 
        #******* RECORRER TODAS LAS PARCIALIDADES CONTANDO LOS ARCHIVOS DE LA SIGUIENTE ESTRUCTURA de las Carpetas en REVISORES  por cada parcialidad.
        #******* 

        parcialidad_0_7_10 = parcialidad[0:7]
        if parcialidad_0_7_10   == '0029-14':
            parcialidad_0_7_10 = parcialidad[0:10]
        elif parcialidad_0_7_10 == '032ESO-':
            parcialidad_0_7_10 = parcialidad[0:9]
        elif parcialidad_0_7_10 == '032ESP-':
            parcialidad_0_7_10 = parcialidad[0:9]
      
        #******* Abrir Planilla CONTROL DOCUMENTOS ING DEF Pxxxx-xx con las 8 hojas para traspasar a BAT
        archivo_parcialidad = ruta_base + parcialidad + '\\CONTROL DOCUMENTOS AS-BUILT P' + parcialidad_0_7_10 + '.xlsx'
 
        if not os.path.exists(archivo_parcialidad):
              log_file.write(f'Parcialidad AS BUILT: {parcialidad_0_7_10} SIN ARCHIVO AS BUILT {archivo_parcialidad}\n')
        else:
                print(f'Procesando Parcialidad AS BUILT: {parcialidad_0_7_10} ARCHIVO:  {archivo_parcialidad}')

                #******* Cargar HOJA: ÚLTIMA VERSIÓN del archivo Excel en un DataFrame (DF_xxxx)

                #ASBUILT

                #ACTUALIZA DOC VIG
                #ACTUALIZA REV LETRA PARCI APRO
                #ACTUALIZA REV NUM PARCI
                

                workbook = openpyxl.load_workbook(archivo_parcialidad, data_only=True, read_only=True)
                
                # Lee el archivo Excel para obtener los nombres de las hojas
                xl = pd.ExcelFile(archivo_parcialidad)
                nombres_hojas = xl.sheet_names
                valor_uvAB = ' '

                if not 'ÚLTIMA VERSIÓN' in nombres_hojas:
                        print(f'La hoja ASBUILT ÚLTIMA VERSIÓN no existe en {archivo_excel}.')  
                        log_file.write(f'La hoja ASBUILT ÚLTIMA VERSIÓN no existe en {archivo_excel}\n')
                else: 
                        sheet_ultima_version = workbook['ÚLTIMA VERSIÓN']
                        # Accede al valor de las celda 'Z3'
                        valor_uvAB = sheet_ultima_version['Z3'].value

                ruta_subdirectorio = os.path.join(ruta_base, parcialidad)
                # Iterar a través de la estructura y crear carpetas si no existen
                for carpeta in estructura:
                    ruta_carpeta = os.path.join(ruta_subdirectorio, carpeta)
                    elementos = os.listdir(ruta_carpeta)
                    archivos = [elemento for elemento in elementos if os.path.isfile(os.path.join(ruta_carpeta, elemento))]
                    cantidad_de_archivos = len(archivos)
                    
                    print(f'{parcialidad_0_7_10:20}\t{carpeta:50}\t{cantidad_de_archivos}\t (uvAB:{valor_uvAB})')             
                    log_file.write(f'{parcialidad_0_7_10:20}\t{carpeta:50}\t{cantidad_de_archivos}\t (uvAB:{valor_uvAB})\n')


                log_file.write(f'\n')
                    
print("Contabilizacion finalizada. Los resultados se han guardado en R:\01 PARCIALIDADES\0000-00 ADMINISTRACION\LOG en el archivo de log_CuadraturaProvisoriaAsBuilt.")
log_file.close

