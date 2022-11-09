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

#coversion de string para las fechas %Y-%m-%d %H:%M  en vez de %d/%m/%Y %H:%M
#dt_objeto es la variable creada para la conversion de la fecha a formato datetime.time
dt_objeto = datetime.strptime(fecha, "%Y-%m-%d %H:%M")

dt_resta = datetime.today() - dt_objeto
print (dt_resta)