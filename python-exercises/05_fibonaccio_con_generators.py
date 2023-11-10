
def fibonacci(num):
    anterior = 1
    actual=0
    siguiente=0
    for i in range(num):
        yield siguiente
        siguiente=anterior+actual
        anterior = actual
        actual = siguiente

for x in fibonacci(100):
    print(x, end=" ")

