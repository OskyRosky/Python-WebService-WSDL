#  Exportar el tiempo en un .txt
# C:/Users/oscar/Desktop/WSDL del Ministerio Mario/scheduled_code.py

from datetime import datetime 

with open('timestamps.txt', 'a+') as file: 
    file.write(str(datetime.now()))