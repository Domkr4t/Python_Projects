from functools import reduce

q = int(input('Факториал какого числа вы хотите получить? Число: '))

try:
    def factorial1(i):
        reduce(lambda x, y: x * y, [i for i in range(1, q + 1)])


    print(factorial1(q))
except Exception:
    print("Факториал не определён")


# ИЛИ


def factorial(n):
    assert n >= 0, 'Факториал не определён'
    if n == 0:
        return 1
    return factorial(n - 1) * n


print(factorial(5))
