import pyshorteners

s = pyshorteners.Shortener()

url = input('Введите вашу ссылку для её сокращения: ')
short_url = s.tinyurl.short(f'{url}')

print(f'Сокращенная ссылка = {short_url}')