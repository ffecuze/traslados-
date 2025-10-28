import openpyxl as opp
import pandas as pd
import numpy as np
import win32com.client
import fitz
import os
import datetime as dt
from PIL import Image
 
informacion = {
    'transportista': ['ERICK MONTENEGRO', 'PDD3279'],
    'fecha': dt.datetime.today().strftime('%Y-%m-%d'),
    'SM': ['AZAMA', 'JOSE LUNA'],
    'MN': ['MANUELA','CAMILO GUZMAN'],
    'MB': ['MARIA BONITA', 'ORLANDO FARINANGO'],
    'CW': ['FLORES DE LA MONTAÑA','EDWIN COBOS'],
    'FS': ['SANTA MONICA','JAIME CARDENAS'],
    'Flopack': ['Flopack',"Nelson Panchana"],
    'Denmar': ['DENMAR',"JAVIER CUZCO"]
}
 
nombre_archivo = os.path.abspath("C:\\Excel_to_PDF\\FormatoTraslados.xlsx")
path = os.path.split(nombre_archivo)
 
df = pd.read_excel("TransporteInterno2024.xlsx", sheet_name='DATOS')
f_df = df[np.logical_and(df['FECHA'] == informacion['fecha'], df['Preguntar en:'] == 'Bodega')]
 
# Crear mapeo dinámico de aprobadores por área solicitante
aprobadores_por_area = f_df[['Area solicitante', 'Aprobador']].drop_duplicates().set_index('Area solicitante')['Aprobador'].to_dict()
 
informacion['up-origen'] = f_df['UP-ORIGEN'].unique()
informacion['up-destino'] = f_df['UP-DESTINO'].unique()


for up_origen in informacion['up-origen']:
    for up_destino in informacion['up-destino']:
        if up_origen == up_destino:
            continue
 
        df2 = f_df[np.logical_and(f_df['UP-ORIGEN'] == up_origen, f_df['UP-DESTINO'] == up_destino)]
        if len(df2) == 0:
            continue
        elif len(df2) > 13:
            print("Error: revisar la cantidad de artículos que van a ser trasladados y notificar si hay cambio de formato")
            break
 
        wb = opp.load_workbook(filename=nombre_archivo)
        sheet = wb['Traslado']
        for row in sheet['C11':'H23']:
            for cell in row:
                cell.value = ""
 
        transportista = informacion['transportista'][0]
        placa = informacion['transportista'][1]
        f_envia = informacion[up_origen][0]
        jefe_almacen_envia = informacion[up_origen][1]
        f_solicita = informacion[up_destino][0]
        jefe_almacen_solicita = informacion[up_destino][1]
 
        # Obtener área y aprobador dinámicamente del Excel
        area_solicitante = df2['Area solicitante'].iloc[0]
        aprobador = aprobadores_por_area.get(area_solicitante, "NO DEFINIDO")
 
        sheet['G4'] = transportista
        sheet['G6'] = placa
        sheet['G8'] = aprobador
        sheet['G2'] = f_envia
        sheet['D2'] = f_solicita
        sheet['D4'] = jefe_almacen_solicita
        sheet['D6'] = area_solicitante
        sheet['C35'] = f"NOMBRE: {jefe_almacen_envia}"
 
        celda_inicio = 11
        for i in range(len(df2)):
            sheet[f'C{celda_inicio}'] = df2.iloc[i, 1]
            sheet[f'D{celda_inicio}'] = df2.iloc[i, 4]
            sheet[f'E{celda_inicio}'] = df2.iloc[i, 5]
            sheet[f'F{celda_inicio}'] = df2.iloc[i, 7]
            sheet[f'G{celda_inicio}'] = df2.iloc[i, 6]
            sheet[f'H{celda_inicio}'] = df2.iloc[i, 12]
            celda_inicio += 1
 
        wb.save(filename=nombre_archivo)
        wb.close()
 
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        sheets = excel.Workbooks.Open(nombre_archivo)
        name_pdf = f"TRASLADO_{up_origen}-{up_destino}.pdf"
        sheets.ExportAsFixedFormat(0, os.path.join(path[0], name_pdf))
        sheets.Close()
        excel.Quit()
 
        doc = fitz.open(os.path.join(path[0], name_pdf))
        page = doc[0]
        rect = fitz.Rect(40, 320, 300, 400)
        page.insert_image(rect, filename="SUMILLA-removebg-preview.png")
        name2 = f"TRASLADO {up_origen}-{up_destino}-signed.pdf"
        doc.save(filename=os.path.join(path[0], name2))





  