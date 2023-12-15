import openpyxl
import os
import warnings
import tkinter as tk
import pandas as pd

folder_path = 'C:\\Users\\alan.riquelmes\\Desktop\\DESCARGAS'
# Función que se ejecuta al hacer clic en el botón 'Ejecutar'
def run_code():
    folder_path = folder_path_var.get()
    if folder_path:
        dataframe = []

dataframe = []

#Lee cada archivo que está en la carpeta
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)
        sheets_dict = pd.read_excel(file_path, sheet_name=None)
        for name, sheet in sheets_dict.items():
            sheet['filename'] = filename
            sheet['sheetname'] = name
            dataframe.append(sheet)

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

df0 = dataframe[0]
df1 = dataframe[1]
df2 = dataframe[2]
df3 = dataframe[3]
df4 = dataframe[4]
df5 = dataframe[5]

#Estructura de la tabla
 #FACTURA - CHASIS - IMPORTE - FECHA DE VENCIMIENTO - FECHA FACTURA - NRO DE OPERACIÓN - EMPRESA
#Agrego en cada columna del dataframe una columna con el nombre de la empresa

#Con esto lo que hago es seleccionar las columnas que me interesa extraer

df0_selected = df0.loc[:, ["Factura.1", "Chasis", "Importe" ,"Fin", "F. franquicia"]]
df1_selected = df1.loc[:, ["Factura", "Chasis", "Importe" ,"Fin", "F. franquicia"]]
df2_selected = df2.loc[:, ["Factura","Chasis" , "Imp.a Pagar" , "Fch.Vto. Free", "Fch.Factura"]]
df3_selected = df3.loc[:, ["Factura","Chasis" , "Imp.a Pagar" , "Fch.Vto. Free","Fch.Factura"]]
df4_selected = df4.loc[:, ["Unnamed: 5","Unnamed: 6" , "Unnamed: 8" , "Unnamed: 11", "Unnamed: 10","Unnamed: 4"]]
df5_selected = df5.loc[:, ["Unnamed: 6","Unnamed: 7" , "Unnamed: 9" , "Unnamed: 13", "Unnamed: 12","Unnamed: 5"]]

df4_selected = df4_selected.drop([0,1,2])
df5_selected = df5_selected.drop([0,1,2])

# Corrección para df0_selected
df0_selected["Fin"] = pd.to_datetime(df0_selected["Fin"]).dt.strftime("%d/%m/%Y")
df0_selected["F. franquicia"] = pd.to_datetime(df0_selected["F. franquicia"]).dt.strftime("%d/%m/%Y")

# Corrección para df1_selected
df1_selected["Fin"] = pd.to_datetime(df1_selected["Fin"]).dt.strftime("%d/%m/%Y")
df1_selected["F. franquicia"] = pd.to_datetime(df1_selected["F. franquicia"]).dt.strftime("%d/%m/%Y")

# Corrección para df2_selected
df2_selected["Fch.Vto. Free"] = pd.to_datetime(df2_selected["Fch.Vto. Free"]).dt.strftime("%d/%m/%Y")
df2_selected["Fch.Factura"] = pd.to_datetime(df2_selected["Fch.Factura"]).dt.strftime("%d/%m/%Y")

# Corrección para df3_selected
df3_selected["Fch.Vto. Free"] = pd.to_datetime(df3_selected["Fch.Vto. Free"]).dt.strftime("%d/%m/%Y")
df3_selected["Fch.Factura"] = pd.to_datetime(df3_selected["Fch.Factura"]).dt.strftime("%d/%m/%Y")

#Le doy el nombre a las columnas de acuerdo a la estructura de la tabla

df0_selected = df0_selected.rename(columns={'Factura.1': 'FACTURA', 'Chasis': 'CHASIS', 'Importe': 'IMPORTE', 'Fin': 'FECHA DE VENCIMIENTO','F. franquicia': 'Fecha Factura'})
df1_selected = df1_selected.rename(columns={'Factura': 'FACTURA', 'Chasis': 'CHASIS', 'Importe': 'IMPORTE', 'Fin': 'FECHA DE VENCIMIENTO','F. franquicia': 'Fecha Factura'})
df2_selected = df2_selected.rename(columns={'Factura': 'FACTURA', 'Chasis': 'CHASIS', 'Imp.a Pagar': 'IMPORTE', 'Fch.Vto. Free': 'FECHA DE VENCIMIENTO','Fch.Factura':'Fecha Factura'})
df3_selected = df3_selected.rename(columns={'Factura': 'FACTURA', 'Chasis': 'CHASIS', 'Imp.a Pagar': 'IMPORTE', 'Fch.Vto. Free': 'FECHA DE VENCIMIENTO','Fch.Factura':'Fecha Factura' })
df4_selected = df4_selected.rename(columns={'Unnamed: 5': 'FACTURA', 'Unnamed: 6': 'CHASIS', 'Unnamed: 8': 'IMPORTE', 'Unnamed: 11': 'FECHA DE VENCIMIENTO','Unnamed: 10' : 'Fecha Factura','Unnamed: 4' :'Nro de operacion'})
df5_selected = df5_selected.rename(columns={'Unnamed: 6': 'FACTURA', 'Unnamed: 7': 'CHASIS', 'Unnamed: 9': 'IMPORTE', 'Unnamed: 13': 'FECHA DE VENCIMIENTO','Unnamed: 12': 'Fecha Factura','Unnamed: 5' :'Nro de operacion'})
df0_selected["NOMBRE EMPRESA"] = "NIX"
df1_selected["NOMBRE EMPRESA"] = "TAGLE"
df2_selected["NOMBRE EMPRESA"] = "MOTCOR"
df3_selected["NOMBRE EMPRESA"] = "RUBIC"
df4_selected["NOMBRE EMPRESA"] = "AVANT"
df5_selected["NOMBRE EMPRESA"] = "ROLEN"

df_concatenado = pd.concat([df0_selected, df1_selected, df2_selected , df3_selected , df4_selected, df5_selected])

with pd.ExcelWriter('consolidado.xlsx') as writer:
    #Pegar los dataframes en hojas diferentes del archivo Excel
    df_concatenado.to_excel(writer, sheet_name="Hoja1", index=False)
