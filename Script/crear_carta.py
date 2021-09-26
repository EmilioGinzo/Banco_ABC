import pandas as pd
from docxtpl import DocxTemplate
from pandas.core.frame import DataFrame

#recibe el nombre de un cliente y retorna un diccionario con el nombre del cliente, el numero de cuenta y el saldo y la suma de todos los saldos
def datos_cliente(nombre_cliente: str, df: DataFrame) -> object:
    cuenta = df.loc[df.Cliente == nombre_cliente, ["N de Cuenta"]].to_string(index = False, header = False)
    saldo = df.loc[df.Cliente == nombre_cliente, ["Saldo"]].to_string(index = False, header = False)
    contenido = {
        'nombre_cliente': nombre_cliente,
        'numero_de_cuenta': cuenta,
        'saldo': saldo,}
    return contenido 

def crear_docx(nombre_cliente: str, df: DataFrame) -> None:
    print("creando documentos de " + nombre_cliente)
    doc = DocxTemplate("..\Output\BANCO ABC - Carta.docx")

    contenido = datos_cliente(nombre_cliente, df)
    doc.render(contenido)
    doc.save("..\Output\BANCO ABC - Carta " + nombre_cliente + ".docx")


def cuentas_clientes(df: DataFrame) -> None:
    prev_cliente = ""
    for cliente in df.Cliente:
        if prev_cliente != cliente:
            crear_docx(cliente, df)
            prev_cliente = cliente

if __name__ == "__main__":
    df = pd.read_excel("..\Input\Datos Clientes.xlsx",sheet_name = "Sheet1", header = 0)
    cuentas_clientes(df)