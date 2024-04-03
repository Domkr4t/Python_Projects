import qrcode

link = 'https://qoogle.com'

file_name = f'QRcode_of_{link.split("/")[-1]}.png'

img = qrcode.make(link)

img.save(file_name)