import random
import sys

if len(sys.argv)==3:
    inicio = int(sys.argv[1])
    final = int(sys.argv[2])
else:
    inicio = 1
    final= 3

random_number = random.randint(inicio, final)


def guess_function(inicio, final, guess, random_number):
    if inicio <= guess <= final:
        if guess == random_number:
            print("¡Muy bien!")
            return True
        elif guess != random_number:
            print("NOP")
    else:
        print("Número fuera de rango")
        return False

if __name__ == '__main__':
    while True:
        try: 
            guess = int(input(f"Adivina el número del {inicio} al {final}: "))
            if guess_function(inicio, final, guess, random_number):
                break
        except ValueError:
            print("Tiene que ser un número entero!")
