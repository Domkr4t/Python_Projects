import requests
from bs4 import BeautifulSoup


#данная функция отправляет запрос на сайт, и возвращает HTML-код страницы
def get_html(url):
    response = requests.get(url)
    return response.text

#данная фукнция парсит HTML-код
def get_data(html):
    soup = BeautifulSoup(html, "lxml") #первый параметр - это сам HTML-код, второй параметр - вид парсера(лучше выбирать как раз lxml)
    title = soup.find('p', {'class': 'site-title'}).text
    return title


def main():
    url = 'https://ru.wordpress.org/'
    print(get_data(get_html(url)))



if __name__ == '__main__':
    main()
