# -*- coding: utf-8 -*-
"""
Created on Wed May 18 14:54:02 2022

@author: Sebastian
"""

import pandas as pd
from sqlalchemy import create_engine
import numpy as np
import win32com.client
import conexion

listado = 'rposada@celsia.com;hbriceno@celsia.com;mmoya@celsia.com;dgomezg@celsia.com;dffigueroa@celsia.com'
listado_sectores = 'Jorge Olmedo Osorio Salgado <joosorio@celsia.com>; Walter Luis Wilches <wwilches@celsia.com>; Eduardo Andres Burbano Bolaños <eburbano@celsia.com>; John Jairo Guzman Belalcazar <jjguzman@celsia.com>; Rosmer Parra Vidal <rparra@celsia.com>; Raul Eduardo Riascos Castro <rriascos@celsia.com>; Arbey Murillo Mosquera <amurillo@celsia.com>'
path_no_regulados = 'D:\\PRIME\\ANALISIS_CLIENTES_NO_REGULADOS.xlsx'
path_peajes = 'D:\\PRIME\\ANALISIS_CLIENTES_PEAJES.xlsx'


"""
==================================================
CONEXION BD
==================================================
"""
con = conexion.con
cursor = con.cursor()

"""
==================================================
CONSULTA SQL
==================================================
"""
data = pd.read_sql_query(conexion.sql, con)
data = data.fillna(0)
columnas = data.columns.tolist()

columna_erase = (columnas[len(columnas)-1:len(columnas)])
for i in columna_erase:
    del data[i]

columnas = data.columns.tolist()
data['PROMEDIO'] = data.loc[:, '3-2022':'8-2022'].mean(axis=1) # Calculo el promedio
data['REDUCCION'] = (data['9-2022'] / data['PROMEDIO']) * 100
data = data.fillna(0)
data_list = data.values.tolist()
data_revisar = data[(data['REDUCCION'] <= 50) & (data['REDUCCION']
                                                 > 0) & (data['PROMEDIO'] > 10000)].sort_values('PROMEDIO', ascending=False)
data_revisar.to_excel(
    path_no_regulados, index=False)


def envios_analisis_fronteras(to, copia, adjunto):
    outlook = win32com.client.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    html_mesagge = open('no_regulado.html').read()
    email.Attachments.Add(adjunto)
    email.To = to
    email.CC = copia
    email.Subject = 'ANALISIS FRONTERAS NO REGULADAS VALLE NORTE Y VALLE SUR - SEPTIEMBRE 2022'
    #email.Subject = f'ANALISIS FRONTERAS PEAJE VALLE NORTE Y VALLE SUR'
    email.HTMLBody = html_mesagge
    email.Display(False)
    email.Send()
    print('ok')


#envios_analisis_fronteras(listado_sectores, listado, path_no_regulados)
