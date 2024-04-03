import requests
from bs4 import BeautifulSoup
import csv


#данная функция отправляет запрос на сайт, и возвращает HTML-код страницы
def get_html(url):
    response = requests.get(url)
    return response.text


def write_csv(data):
    with open('coinmarket.csv', 'a') as f:
        writer = csv.writer(f)

        writer.writerow([data['name'], data['source'], data['price']])


#данная фукнция парсит HTML-код
def get_page_data(html):
    soup = BeautifulSoup(html, "lxml")

    trs = soup.find('table', {'class':'sc-f7a61dda-3 kCSmOD cmc-table'}).find('tbody')

    for i in trs:
        tds = i.find_all('td')
        name = tds[2].find('a', class_='cmc-link').get('href')[12:].replace('/', '').replace('-', ' ').title()
        price = tds[3].text.replace(',', ' ')
        url = tds[2].find('a', class_='cmc-link').get('href')[1:]
        url_end = f'https://coinmarketcap.com/{url}'

        data = {'name': name, 'source': url_end, 'price': price}

        write_csv(data)



def main():
    url = 'https://coinmarketcap.com/'
    get_page_data(get_html(url))


if __name__ == '__main__':
    main()
