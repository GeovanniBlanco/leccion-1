nombre = input("ingresa el nombre del libro: ")
id = input("ingresa el ID del libro ")
precio = input("ingresa el precio del libro: ")
envio = input("envio gratis (true or false: ")

print(f"el nombre del libro es {nombre}")
print(f"el ID del libro elegido es {int(id)}")
print(f"el precio del libro es {float(precio)}")
print(f"el envio gratuito sera {bool(envio)}")