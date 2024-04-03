from stegano import lsb  # не подходит для русского текста
from stegano import exifHeader # подходит для русского текста
from steganocryptopy.steganography import Steganography # для генерации секретного ключа для дешифрования текста (только png)


#1 параметр - путь до файла в котором требуется сохранить ключ
Steganography.generate_key('key.txt')

#1 аргумент - путь для ключа, 2 - путь до изображения в которое закодировать, 3 - путь до файла с текстом
secret_with_key_only_png = Steganography.encrypt('key.txt', 'pngwing.com.png', 'text-for-encrypt.txt')
#аргумент - имя нового файла в котором будет сокрыт текст или путь
secret_with_key_only_png.save('vk-password.png')

#1 аргумент - путь для ключа, 2 - путь к файлу в котором сокрыт текст
result_with_key = Steganography.decrypt('key.txt', 'vk-password.png')
print(result_with_key)



#первый аргумент - путь до файла в котором требуется скрыть текст, второй - сам текст
secret_in_png = lsb.hide('pngwing.com.png.png', 'From binance')
#аргумент - имя нового файла в котором будет сокрыт текст или путь
secret_in_png.save('sadasdasd.png')

#чтение сокрытого текста, аргумент - путь к файлу в котором сокрыт текст
result_png = lsb.reveal('asdasdasdas.png')
print(result_png)


#первый аргумент - путь до файла в котором требуется скрыть текст, второй - путь к файлу в котором сокрыт текст, третий - сам текст
secret_in_jpg = exifHeader.hide('krasivye-zhivotnye-v-dikoj-prirode.jpg', 'Wifi.jpg', 'Пароль от вайфая')

result_jpg = exifHeader.reveal('Wifi.jpg')
print(result_jpg.decode())