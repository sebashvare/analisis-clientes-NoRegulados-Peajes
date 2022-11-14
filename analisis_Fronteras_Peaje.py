# -*- coding: utf-8 -*-
"""
Created on Wed May 18 14:54:02 2022

@author: Sebastian
"""

import pandas as pd
from sqlalchemy import create_engine
import win32com.client

listado = 'rposada@celsia.com;hbriceno@celsia.com;mmoya@celsia.com;dgomezg@celsia.com;dffigueroa@celsia.com'
listado_sectores = 'Jorge Olmedo Osorio Salgado <joosorio@celsia.com>; Walter Luis Wilches <wwilches@celsia.com>; Eduardo Andres Burbano Bolaños <eburbano@celsia.com>; John Jairo Guzman Belalcazar <jjguzman@celsia.com>; Rosmer Parra Vidal <rparra@celsia.com>; Raul Eduardo Riascos Castro <rriascos@celsia.com>; Arbey Murillo Mosquera <amurillo@celsia.com>'
path_no_regulados = 'D:\\PRIME\\ANALISIS_CLIENTES_NO_REGULADOS.xlsx'
path_peajes = 'D:\\PRIME\\ANALISIS_CLIENTES_PEAJES.xlsx'

sql = """
WITH CLIENTES (CODIGOO_SIC, NOMBRE_FRONTERA , IMPO_EXPO,SECTOR,MUNICIPIO, TOTAL , MES2) AS (
SELECT DISTINCT 
	   AE.[CODIGO_SIC]
      ,CODIGO_PROPIO
	  ,[IMPO_EXPO]
      ,MT.NOMBRE_SUCURSAL
	  ,MT.MUNICIPIO
	  ,SUM([TOTAL]) AS TOTAL
      ,CONCAT(MES, '-', ANIO) AS MES2
FROM [INDICADORES].[dbo].[AENC] AS AE
INNER JOIN (SELECT DISTINCT TF.CODIGO_FRONTERA, TF.FACTOR_PERDIDAS, TF.AGENTE_COMERCIAL_QUE_IMPORTA, TF.MERCADO_COMERCIALIZACION_QUE_EXPORTA, NT FROM TFROC AS TF 
			WHERE 
			TF.AGENTE_COMERCIAL_QUE_EXPORTA = 'EPSC' 
		  --AND TF.AGENTE_COMERCIAL_QUE_IMPORTA IN ('EPSC') -- EVALUAR NO REGULADOS
		  AND TF.AGENTE_COMERCIAL_QUE_IMPORTA NOT IN ('EPSC','EPSG') -- EVALUAR PEAJES
			--AND TF.MERCADO_COMERCIALIZACION_QUE_EXPORTA IN ('EPSD') -- METODOLOGIA ANTES DE JULIO-2021
			--AND TF.MERCADO_COMERCIALIZACION_QUE_IMPORTA = 'EPSD' -- METODOLOGIA ANTES DE JULIO-2021
			AND TF.MERCADO_COMERCIALIZACION_QUE_EXPORTA IN ('VACM') --METOLOGIA NUEVA
			AND TF.MERCADO_COMERCIALIZACION_QUE_IMPORTA IN ('VACM')--METOLOGIA NUEVA
			)AS TF
ON (TF.CODIGO_FRONTERA = AE.CODIGO_SIC)
LEFT JOIN [INDICADORES].[dbo].[MITHRA] AS MT
ON (MT.CODIGO_SIC = AE.CODIGO_SIC)
WHERE AE.IMPO_EXPO = 'E'
GROUP BY AE.[CODIGO_SIC]
		 ,CODIGO_PROPIO
		 ,[IMPO_EXPO]
		 ,[MES]
		 ,[ANIO]
		 ,MT.NOMBRE_SUCURSAL
		 ,MT.MUNICIPIO
)

SELECT * FROM CLIENTES
PIVOT (SUM(TOTAL) FOR MES2 IN ([1-2021],
[2-2021],
[3-2021],
[4-2021],
[5-2021],
[6-2021],
[7-2021],
[8-2021],
[9-2021],
[10-2021],
[11-2021],
[12-2021],
[1-2022],
[2-2022],
[3-2022],
[4-2022],
[5-2022],
[6-2022],
[7-2022],
[8-2022],
[9-2022],
[10-2022]
)) PVT

"""
# VARIABLES DE CONEXION
SERVERNAME = 'LOCALHOST\SQLEXPRESS'
DB = 'INDICADORES'
DRIVER = 'ODBC Driver 17 for SQL Server'

# INSTANCIAMOS LA CONEXION CON SLQALCHEMY
engine = create_engine(
    f"mssql+pyodbc://{SERVERNAME}/{DB}?driver={DRIVER}"
)
con = engine.raw_connection()
cursor = con.cursor()


data = pd.read_sql_query(sql, con)
data = data.fillna(0)
columnas = data.columns.tolist()

columna_erase = (columnas[len(columnas)-1:len(columnas)])
for i in columna_erase:
    del data[i]

columnas = data.columns.tolist()
data['PROMEDIO'] = data.loc[:, '3-2022':'8-2022'].mean(axis=1) # --------------> MESES PARA CALCULAR EL PROMEDIO
data['REDUCCION'] = (data['9-2022'] / data['PROMEDIO']) * 100 # --------------> MES PARA EVALUAR LA DISMINUCION DE ENERGIA SEGUN EL PROMEDIO
data = data.fillna(0)
data_list = data.values.tolist()
data_revisar = data[(data['REDUCCION'] <= 50) & (data['REDUCCION']
                                                              > 0) & (data['PROMEDIO'] > 10000)].sort_values('PROMEDIO', ascending=False)
data_revisar.to_excel(
    path_peajes, index=False)


def envios_analisis_fronteras(to, copia, adjunto):
    outlook = win32com.client.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    html_mesagge = open('peaje.html').read()
    email.Attachments.Add(adjunto)
    email.To = to
    email.CC = copia
    email.Subject = 'ANALISIS FRONTERAS PEAJE VALLE NORTE Y VALLE SUR - SEPTIEMBRE 2022'
    #email.Subject = f'ANALISIS FRONTERAS PEAJE VALLE NORTE Y VALLE SUR'
    email.HTMLBody = html_mesagge
    email.Display(False)
    email.Send()
    print('ok')

envios_analisis_fronteras(listado_sectores,listado, path_peajes)
