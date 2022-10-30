############

from datetime import datetime
start=datetime.now()


#######################
#        Modulos      #
#######################
 
import pandas as pd
from zeep.helpers import serialize_object
from requests import Session
from requests.auth import HTTPBasicAuth
from zeep import Client
from zeep.transports import Transport
# from zeep import xsd

# Requerimientos
# pip install pandas
# pip install openpyxl
# pip install zeep

user = 'Ws_cgr'
password = 'Claves+2021'
wsdl = 'https://sigaf.hacienda.go.cr/sap/bc/srt/wsdl/flv_10002A111AD1/bndg_url/sap/bc/srt/rfc/sap/zfmg_reports/500/zfmg_reports/binding?sap-client=500'
endpoint = 'https://sigaf.hacienda.go.cr/sap/bc/srt/rfc/sap/zfmg_reports/500/zfmg_reports/binding'


def conectar_ws_hacienda():

    try:

        # Se conecta utilizando autenticacion basica
        
        session = Session()
        session.auth = HTTPBasicAuth(user, password)
        client = Client(wsdl, transport=Transport(session=session))

        # Se conecta al WS mediante el Endpoint
        
        service = client.create_service(
            '{urn:sap-com:document:sap:soap:functions:mc-style}binding_soap12', endpoint)

###############################
# Consumos y parserisación    #
################################ 

# Gastos acumulados 

        Entidadcp = {'item': ['POWR','PEJC']}


        datos1 = service.ZwsYMhd76000042('101', '40199999', 2021, Entidadcp, 1, 12)

        # Serializa los datos convertirlos a un formato usable.
        datos1 = serialize_object(datos1)

        # Transforma los datos en un Dataframe de Pandas
        df1 = pd.DataFrame(datos1)

# Gastos mensuales  

        Entidadcp = {'item': ['POWR','PEJC']}

        
        datos2 = service.ZwsZfmMensual(
            '101', '40199999', 2021, Entidadcp)

        # Serializa los datos convertirlos a un formato usable.
        datos2 = serialize_object(datos2)

        # Transforma los datos en un Dataframe de Pandas
        df2 = pd.DataFrame(datos2)


# Ingresos 

        # Ejecuta la funcion ZwsZinforme37New, enviando los parametros.
        datos3 = service.ZwsZinforme37New(
                                           '101', '40199999', 2021, 'PEJC', 1, 11)

        # Serializa los datos convertirlos a un formato usable.
        datos3 = serialize_object(datos3)

        # Transforma los datos en un Dataframe de Pandas
        df3 = pd.DataFrame(datos3)

############################################
#    Exporta los datos al formato deseado  #
############################################
        
        exportar(df1, 'excel', 'gastos_acumulados.xlsx')
        
        exportar(df2, 'excel', 'gastos_mensuales.xlsx')

        exportar(df3, 'excel', 'ingresos.xlsx')        

    except Exception as e:
        print(e)

# Exportación de los gastos acumulados

def exportar(df1, tipo, nombre):
    try:
        if tipo == 'excel':
            writer = pd.ExcelWriter(nombre)
            df1.to_excel(writer,'Datos')
            writer.save()
            print('Archivo exportado a Excel:', nombre)
            return
        if tipo == 'csv':
            df1.to_csv('archivo.csv', sep=';', index=False)
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

print (datetime.now()-start)

