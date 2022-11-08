#importacion de librerias y manejo de fechas

from datetime import date, time,  datetime
datetime.today()

print(date.today())

dt = datetime.now()
dt.year
dt.month
dt.day
dt.hour
dt.minute
dt.second
dt.microsecond

print(dt)

fecha = input('ingresa una fecha: ')

#coversion de string para las fechas
dt_objeto = datetime.strptime(fecha, "%d/%m/%Y %H:%M:%S")


print (dt_objeto)