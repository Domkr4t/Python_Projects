import requests
import fake_useragent
from bs4 import BeautifulSoup


user = fake_useragent.UserAgent().random

# header = {'user-agent':
#         'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
#         'Chrome/94.0.4606.85 YaBrowser/21.11.3.927 Yowser/2.5 Safari/537.36'}

header = {'user-agent': user} #если подменять с помощью fake_useragent, лучше делать так
#нужен чтобы браузер не понимал что это парсер, находится в главном запросе в F12, в заголовках в конце

link = "https://browser-info.ru/"

response = requests.get(link, headers = header).text
#requests обычное использование, sesion - если требуется сохранять объекты для входа например, обязаетльно нужно создать объект сессии
# .get для того чтобы чисто брать данные со страницы, requests.post(link) #post если нужно что то еще передать, например логин и пароль
# .text для получения HTML, .content для получения байтов


soup = BeautifulSoup(response, 'lxml')
block = soup.find('div', id = "tool_padding")   #find для поиска одного блока, find_all для поиска множества блоков


js = block.find('div', id = "javascript_check").find_all('span')[1].text
flash = block.find('div', id = "flash_version").find_all('span')[1].text
user_agent = block.find('div', id = "user_agent").text

print(f'Javascript {js[:-1]}\nFlash {flash}\n{user_agent}')
