a = ' OpenMe.txt'
b = 0
text = [x for x in range(10)]  # количество чисел в файле
for i in range(10):   # количество созданных файлов
    c = list(a)
    c[0] = b
    b += 1
    w = "".join(str(x) for x in c)
    with open('{}'.format(w), 'a+') as file:
        file.write(str(text))
