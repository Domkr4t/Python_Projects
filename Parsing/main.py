import requests
import fake_useragent
from bs4 import BeautifulSoup


session = requests.session()
#нужно для сохранения "прогресса", например сохранить данные для входа

data = {
    'inmembername': 'login',
    'inpassword': 'password'
}
#data для хранения пароля и логина


link = "http://forum.ru-board.com/"

response = session.post(link, headers = header, data = data).text
#requests обычное использование, sesion - если требуется сохранять объекты для входа например, обязаетльно нужно создать объект сессии
# .get для того чтобы чисто брать данные со страницы, requests.post(link) #post если нужно что то еще передать, например логин и пароль
# .text для получения HTML, .content для получения байтов


user = fake_useragent.UserAgent().random

data = {
    'inmembername': 'login',
    'inpassword': 'password'
}

header = {'user-agent': user}


response = session.post(link, headers = header, data = data).text

profile_link = 'http://forum.ru-board.com/profile.cgi'
profile_nick = session.get(profile_link, headers = header).find('td').find('b').text

print(profile_nick)
