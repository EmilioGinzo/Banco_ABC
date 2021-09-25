import pandas as pd

df = pd.read_excel("..\Input\Datos Clientes.xlsx",sheet_name = "Sheet1", header = 0)

#imprime los datos (numero de cuenta, saldo, suma de saldo )del cliente que se le pase
def datos_cliente(nombre_cliente: str) -> None:
    print("datos del cliente ", nombre_cliente)
    print(df.loc[df.Cliente == nombre_cliente, ["N de Cuenta", "Saldo"]].to_string(index = False, header = True))
    suma = df.loc[df.Cliente == nombre_cliente, ["Saldo"]].sum()
    print("\nTotal", suma)

#imprime los datos separados de cada cliente
def cuentas_clientes(clientes: object) -> None:
    prev_cliente = ""
    for cliente in clientes:
        if prev_cliente != cliente:
            datos_cliente(cliente)
            prev_cliente = cliente

cuentas_clientes(df.Cliente)