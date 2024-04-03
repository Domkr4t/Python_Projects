def home_player():

    goals_previous_match = input('\nВведите количество мячей хозяев в 5 прошлых матчах через пробел (например 75 67 32 23 68):').split()
    sum_previous_match = 0
    for i in goals_previous_match:
        sum_previous_match += int(i)
    result_goals_previous_matchs = sum_previous_match / 5

    win_lose_previous_match = input('Введите количество побед и поражений хозяев в 5 прошлых матчах  через пробел (например 2 3):').split()
    result_win_lose = int(win_lose_previous_match[0]) * 1 + int(win_lose_previous_match[1]) * (-1.5) + 3

    result_previous_matchs = result_goals_previous_matchs + int(result_win_lose)

    goals_confrontation_match = input('Введите количество мячей хозяев в 5 прошлых очных матчах между этими командами через пробел (например 75 67 32 23 68):').split()
    sum_confrontation_match = 0
    for i in goals_confrontation_match:
        sum_confrontation_match += int(i)
    result_goals_confrontation_matchs = sum_confrontation_match / 5

    result = (result_previous_matchs + result_goals_confrontation_matchs) / 2

    return result

def guest_player():

    goals_previous_match = input('\nВведите количество мячей гостей в 5 прошлых матчах через пробел (например 75 67 32 23 68):').split()
    sum_previous_match = 0
    for i in goals_previous_match:
        sum_previous_match += int(i)
    result_goals_previous_matchs = sum_previous_match / 5

    win_lose_previous_match = input('Введите количество побед и поражений гостей в 5 прошлых матчах  через пробел (например 2 3):').split()
    result_win_lose = int(win_lose_previous_match[0]) * 1 + int(win_lose_previous_match[1]) * (-1.5)

    result_previous_matchs = result_goals_previous_matchs + int(result_win_lose)

    goals_confrontation_match = input('Введите количество мячей гостей в 5 прошлых очных матчах между этими командами через пробел (например 75 67 32 23 68):').split()
    sum_confrontation_match = 0
    for i in goals_confrontation_match:
        sum_confrontation_match += int(i)
    result_goals_confrontation_matchs = sum_confrontation_match / 5

    result = (result_previous_matchs + result_goals_confrontation_matchs) / 2

    return result

def main():
    result_home = home_player()
    result_guest = guest_player()
    print(f'\nСредний тотал команды хозяев: {result_home}')
    print(f'Средний тотал команды гостей: {result_guest}')
    print(f'Общий средний тотал команд: {result_home + result_guest} \n')


while True:
    main()
    if input("Начать анализировать следующий матч? (да/нет) ").strip().upper() != 'ДА':
        break
