from unittest import result
import requests
import fake_useragent
from bs4 import BeautifulSoup
from datetime import datetime, timedelta

def get_goroskop_today(sign):

    user = fake_useragent.UserAgent().random
    header = {'user-agent': user}
    sign.lower()
    if sign == "овен":
        link = "https://horo.mail.ru/prediction/aries/today/"
    if sign == "водолей":
        link = "https://horo.mail.ru/prediction/aquarius/today/"
    response = requests.get(link, headers = header).text



    soup = BeautifulSoup(response, 'lxml')

    goroskop = soup.find('div', class_="article__item article__item_alignment_left article__item_html").text
    buisness = soup.find('span', class_="icon icon_chart icon_score").next_element.text
    love = soup.find('span', class_="icon icon_heart icon_score").next_element.text
    numb_of_day = soup.find('span', class_="icon icon_peak icon_score").next_element.text
    date_today = datetime.now().date()

    result = f"{sign.capitalize()} \nДата: {date_today} \n\n{goroskop} \nБизнес = {buisness}\nЛюбовь = {love}\nЦифра дня = {numb_of_day}\n\n\n\n"

    with open(file=f'result.txt', mode='a+', encoding='utf-8') as file:
        file.write(result)

def get_goroskop_yesterday(sign):

    user = fake_useragent.UserAgent().random
    header = {'user-agent': user}
    sign.lower()
    if sign == "овен":
        link = "https://horo.mail.ru/prediction/aries/yesterday/"
    if sign == "водолей":
        link = "https://horo.mail.ru/prediction/aquarius/yesterday/"
    response = requests.get(link, headers = header).text



    soup = BeautifulSoup(response, 'lxml')

    goroskop = soup.find('div', class_="article__item article__item_alignment_left article__item_html").text
    buisness = soup.find('span', class_="icon icon_chart icon_score").next_element.text
    love = soup.find('span', class_="icon icon_heart icon_score").next_element.text
    numb_of_day = soup.find('span', class_="icon icon_peak icon_score").next_element.text
    date_yesterday = datetime.now().date() - timedelta(days=1)

    result = f"{sign.capitalize()} \nДата: {date_yesterday} \n\n{goroskop} \nБизнес = {buisness}\nЛюбовь = {love}\nЦифра дня = {numb_of_day}\n\n\n\n"

    with open(file=f'result.txt', mode='a+', encoding='utf-8') as file:
        file.write(result)

def get_goroskop_tomorrow(sign):

    user = fake_useragent.UserAgent().random
    header = {'user-agent': user}
    sign.lower()
    if sign == "овен":
        link = "https://horo.mail.ru/prediction/aries/yesterday/"
    if sign == "водолей":
        link = "https://horo.mail.ru/prediction/aquarius/yesterday/"
    response = requests.get(link, headers = header).text



    soup = BeautifulSoup(response, 'lxml')

    goroskop = soup.find('div', class_="article__item article__item_alignment_left article__item_html").text
    buisness = soup.find('span', class_="icon icon_chart icon_score").next_element.text
    love = soup.find('span', class_="icon icon_heart icon_score").next_element.text
    numb_of_day = soup.find('span', class_="icon icon_peak icon_score").next_element.text
    date_tomorrow = datetime.now().date() + timedelta(days=1)

    result = f"{sign.capitalize()} \nДата: {date_tomorrow} \n\n{goroskop} \nБизнес = {buisness}\nЛюбовь = {love}\nЦифра дня = {numb_of_day}\n\n\n\n"

    with open(file=f'result.txt', mode='a+', encoding='utf-8') as file:
        file.write(result)

def get_goroskop_week(sign):

    user = fake_useragent.UserAgent().random
    header = {'user-agent': user}
    sign.lower()
    if sign == "овен":
        link = "https://horo.mail.ru/prediction/aries/week/"
    if sign == "водолей":
        link = "https://horo.mail.ru/prediction/aquarius/week/"
    response = requests.get(link, headers = header).text



    soup = BeautifulSoup(response, 'lxml')

    goroskop = soup.find('div', class_="article__item article__item_alignment_left article__item_html").text

    date = datetime.now().weekday()
    different_of_days = 0

    date1 = datetime.now().weekday()
    different_of_days1 = 6

    while date > 0:
        date -= 1
        different_of_days += 1

    while date1 > 0:
        date1 += 1
        different_of_days1 -= 1

    start_of_week = datetime.now().date() - timedelta(days=different_of_days)
    end_of_week = datetime.now().date() + timedelta(days=different_of_days1)

    result = f"{sign.capitalize()} \nНеделя: от {start_of_week} до {end_of_week}\n\n{goroskop}\n\n\n\n"

    with open(file=f'result.txt', mode='a+', encoding='utf-8') as file:
        file.write(result)

def get_goroskop_year(sign):

    user = fake_useragent.UserAgent().random
    header = {'user-agent': user}
    sign.lower()
    if sign == "овен":
        link = "https://horo.mail.ru/prediction/aries/year/"
    if sign == "водолей":
        link = "https://horo.mail.ru/prediction/aquarius/year/"
    response = requests.get(link, headers = header).text



    soup = BeautifulSoup(response, 'lxml')

    goroskop = soup.find('div', class_="article__item article__item_alignment_left article__item_html").text

    year = datetime.now().year

    result = f"{sign.capitalize()} \nГод: {year}\n\n{goroskop}\n\n\n\n"

    with open(file=f'result.txt', mode='a+', encoding='utf-8') as file:
        file.write(result)


def result_today():
    get_goroskop_today("овен")
    get_goroskop_today("водолей")

def result_yesterday():

    get_goroskop_yesterday("овен")
    get_goroskop_yesterday("водолей")

def result_tomorrow():
    get_goroskop_tomorrow("овен")
    get_goroskop_tomorrow("водолей")

def result_week():
    get_goroskop_week("овен")
    get_goroskop_week("водолей")

def result_year():
    get_goroskop_year("овен")
    get_goroskop_year("водолей")
