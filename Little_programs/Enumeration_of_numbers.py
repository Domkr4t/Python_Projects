def gen_number(number_system: int, len_of_word: int, prefix=None):

    prefix = prefix or []

    if len_of_word == 0:
        print(*prefix, sep='')
        return

    for number in range(number_system):
        prefix.append(number)
        gen_number(number_system, len_of_word - 1, prefix)
        prefix.pop()


gen_number(10, 4)



