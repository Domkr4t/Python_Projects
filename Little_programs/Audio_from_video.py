import moviepy.editor



video_name = input('Введите название видео: ')
video = moviepy.editor.VideoFileClip(f'{video_name}')
audio = video.audio

audio.write_audiofile(f'{video_name.split(".")[0]}_audio.mp3')