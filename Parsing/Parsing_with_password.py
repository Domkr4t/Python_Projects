import requests
import fake_useragent
from bs4 import BeautifulSoup


session = requests.session()
#нужно для сохранения "прогресса", например сохранить данные для входа
user = fake_useragent.UserAgent().random
header = {'user-agent': user} #если подменять с помощью fake_useragent, лучше делать так
#нужен чтобы браузер не понимал что это парсер, находится в главном запросе в F12, в заголовках в конце

data = {'inmembername': 'login', 'inpassword': 'password'}
#data для хранения пароля и логина


link = "http://forum.ru-board.com/misc.cgi"

response = session.post(link, data = data, headers = header).text
#requests обычное использование, sesion - если требуется сохранять объекты для входа например, обязаетльно нужно создать объект сессии
# .get для того чтобы чисто брать данные со страницы, requests.post(link) #post если нужно что то еще передать, например логин и пароль
# .text для получения HTML, .content для получения байтов


profile_link = 'http://forum.ru-board.com/profile.cgi'
profile_nick = session.get(profile_link, headers = header).text

print(profile_nick)
