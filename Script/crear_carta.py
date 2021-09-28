#recibe datos desde un archivo xlsx ubicado en Input y crea una carta para cada cliente basada en el template ubicado en Input
# las cartas ya creadas seran almacenadas en Output

import pandas as pd
from docxtpl import DocxTemplate
from pandas.core.frame import DataFrame
from re import findall
import os

#recibe el nombre de un cliente y el dataframe que contiene al informacion de dicho cliente
# retorna un diccionario con el nombre del cliente, el numero de cuenta, el saldo y la suma de los saldos de sus cuentas
def datos_cliente(nombre_cliente: str, df: DataFrame) -> dict:
    cuenta = df.loc[df.Cliente == nombre_cliente, ["N de Cuenta"]].to_string(index = False, header = False)
    saldo = df.loc[df.Cliente == nombre_cliente, ["Saldo"]].to_string(index = False, header = False)
    saldo_int = findall('[-]?\d+', saldo)
    suma = 0.00
    for monto in saldo_int:
        suma += float(monto)
    
    contenido = {
        'nombre_cliente': nombre_cliente,
        'numero_de_cuenta': cuenta,
        'saldo': saldo,
        'suma': suma}
    return contenido 

#recibe el nombre de un cliente y el df correspondiente al archivo de excel
#crea un documento de tipo docx para el cliente especificado
#el documento creado esta basado en la plantilla que se encuenta en la carpeta Input y sera creado 
#con el siguiente nombre "BANCO ABC - Carta + nombre_cliente + .docx", este documento se encontrara en la carpeta Output
def crear_docx(nombre_cliente: str, df: DataFrame) -> None:
    print("creando documentos de " + nombre_cliente)
    doc = DocxTemplate("..\Input\BANCO ABC - Carta.docx")

    contenido = datos_cliente(nombre_cliente, df)
    doc.render(contenido)
    doc.save("..\Output\BANCO ABC - Carta " + nombre_cliente + ".docx")

#separa las respectivas cuentas de cada cliente y llama a la funcion crear_docx para cada cliente
def cuentas_clientes(df: DataFrame) -> None:
    prev_cliente = ""
    for cliente in df.Cliente:
        if prev_cliente != cliente:
            crear_docx(cliente, df)
            prev_cliente = cliente

if __name__ == "__main__":
    if not os.path.exists('..\Output'):
        os.makedirs('..\Output')
    df = pd.read_excel("..\Input\Datos Clientes.xlsx",sheet_name = "Sheet1", header = 0)
    cuentas_clientes(df)