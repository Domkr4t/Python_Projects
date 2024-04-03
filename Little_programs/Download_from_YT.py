from pytube import YouTube

url = 'https://www.youtube.com/watch?v=ezJ8kYzXe-I'
video = YouTube(url)

print(f'{video.title} is downloading.')

video = video.streams.get_highest_resolution()

video.download()
