import requests
from bs4 import BeautifulSoup
import csv


#данная функция отправляет запрос на сайт, и возвращает HTML-код страницы
def get_html(url):
    response = requests.get(url)
    return response.text

#данная функция обрезает полученные рейтинги (из вида "1 981 общий рейтинг" в "1 981")
def refind(s):
    s.split(' ')
    return s[1:-15].replace(' ', '')

def write_csv(data):
    with open('plugins.csv', 'a') as f:
        writer = csv.writer(f)

        writer.writerow((data['name'],
                        data['source'],
                        data['rating'],
                        data['rating_count'],
                        data['rating_source']))


#данная фукнция парсит HTML-код
def get_data(html):
    soup = BeautifulSoup(html, "lxml") #первый параметр - это сам HTML-код, второй параметр - вид парсера(лучше выбирать как раз lxml)
    popular = soup.find_all('section')

    for i in popular:
        plugins = i.find_all('article')

        for i in plugins:
            name = i.find('h3').text
            source = i.find('h3').find('a').get('href')
            rating = i.find('div', {'class': 'wporg-ratings'}).get('data-rating')
            rating_count = refind(i.find('span', {'class': 'rating-count'}).text)
            rating_source = i.find('span', {'class': 'rating-count'}).find('a').get('href')


            data = {'name': name,
                    'source': source,
                    'rating': rating,
                    'rating_count': rating_count,
                    'rating_source': rating_source}

            write_csv(data)


def main():
    url = 'https://ru.wordpress.org/plugins/'
    get_data(get_html(url))



if __name__ == '__main__':
    main()
