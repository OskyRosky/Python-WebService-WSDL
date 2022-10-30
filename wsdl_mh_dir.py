############

from zeep.transports import Transport
from zeep import Client
from requests.auth import HTTPBasicAuth
from requests import Session
from zeep.helpers import serialize_object
import pandas as pd
from datetime import datetime
start = datetime.now()

#######################
#        Modulos      #
#######################

# from zeep import xsd

# Requerimientos
# pip install pandas
# pip install openpyxl
# pip install zeep

user = 'Ws_cgr'
password = 'Claves+2021'
wsdl = 'https://sigaf.hacienda.go.cr/sap/bc/srt/wsdl/flv_10002A111AD1/bndg_url/sap/bc/srt/rfc/sap/zfmg_reports/500/zfmg_reports/binding?sap-client=500'
endpoint = 'https://sigaf.hacienda.go.cr/sap/bc/srt/rfc/sap/zfmg_reports/500/zfmg_reports/binding'


def obtener_gastos_acumulados(service):
    # Gastos acumulados
    Entidadcp = {'item': ['POWR', 'PEJC']}
    datos = service.ZwsYMhd76000042('101', '40199999', 2021, Entidadcp, 1, 12)

    # Serializa los datos convertirlos a un formato usable.
    datos = serialize_object(datos)

    # Transforma los datos en un Dataframe de Pandas
    df = pd.DataFrame(datos)

    # Exporta a excel
    exportar(df, 'excel', 'C:\\Users\\oscar\\Desktop\\MH_SIGAF\\', 'gastos_acumulados.xlsx')


def obtener_gastos_mensuales(service):
    # Gastos mensuales
    Entidadcp = {'item': ['POWR', 'PEJC']}
    datos = service.ZwsZfmMensual('101', '40199999', 2021, Entidadcp)

    # Serializa los datos convertirlos a un formato usable.
    datos = serialize_object(datos)

    # Transforma los datos en un Dataframe de Pandas
    df = pd.DataFrame(datos)

    # Exporta a excel
    exportar(df, 'excel', 'C:\\Users\\oscar\\Desktop\\MH_SIGAF\\', 'gastos_mensuales.xlsx')


def obtener_ingresos(service):
    # Ingresos

    # Ejecuta la funcion ZwsZinforme37New, enviando los parametros.
    datos = service.ZwsZinforme37New('101', '40199999', 2021, 'PEJC', 1, 11)

    # Serializa los datos convertirlos a un formato usable.
    datos = serialize_object(datos)

    # Transforma los datos en un Dataframe de Pandas
    df = pd.DataFrame(datos)

    # Exporta a excel
    exportar(df, 'excel', 'C:\\Users\\oscar\\Desktop\\MH_SIGAF\\', 'ingresos.xlsx')


def conectar_ws_hacienda():

    try:

        # Se conecta utilizando autenticacion basica
        session = Session()
        session.auth = HTTPBasicAuth(user, password)
        client = Client(wsdl, transport=Transport(session=session))

        # Se conecta al WS mediante el Endpoint
        service = client.create_service(
            '{urn:sap-com:document:sap:soap:functions:mc-style}binding_soap12', endpoint)

    except Exception as e:
        print(e)

    obtener_gastos_acumulados(service)
    obtener_gastos_mensuales(service)
    obtener_ingresos(service)


def exportar(df, tipo, ruta, nombre):
    try:
        if tipo == 'excel':
            writer = pd.ExcelWriter(ruta+nombre)
            df.to_excel(writer, 'Datos')
            writer.save()
            print('Archivo exportado a Excel:', nombre)
            return
        if tipo == 'csv':
            df.to_csv(ruta+nombre, sep=';', index=False)
            print('Archivo exportado a CSV:', nombre)
            return
        else:
            print('Formato incorrecto:')
            return
    except Exception as e:
        print(e)


conectar_ws_hacienda()

###################
#   Tiempo total  #
###################

print(datetime.now()-start)


print("Python " + "3", "11", "0", sep=".")