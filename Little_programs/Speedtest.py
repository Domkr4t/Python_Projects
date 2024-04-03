import speedtest 

test = speedtest.Speedtest(secure=True)
print('Waiting...')

down_speed = test.download()
up_speed = test.upload()

print(f'Download speed: {down_speed / 1024 / 1024:.2f} Mb/s\nUpload speed: {up_speed / 1024 / 1024:.2f} Mb/s')