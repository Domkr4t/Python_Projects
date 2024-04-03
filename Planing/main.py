import schedule
import requests


def greeting():
    """Greeting function"""
    
    todos_dict = {
        '08:00': 'Drink coffee',
        '11:00': 'Work meeting',
        '23:59': 'Hack the Planet!'
    }
    
    print("Day's tasks")
    for k, v in todos_dict.items():
        print(f'{k} - {v}')
        
    response = requests.get(url='https://yobit.net/api/3/ticker/btc_usd')
    data = response.json()
    btc_price = f"BTC: {round(data.get('btc_usd').get('last'), 2)}$\n"
    
    print(btc_price)
        
        
def main():
    
    schedule.every(4).seconds.do(greeting) # выполнять функцию greeting каждые 4 секунды
    schedule.every(5).minutes.do(greeting) # выполнять функцию greeting каждые 5 минут
    schedule.every().hour.do(greeting) # выполнять функцию greeting каждый час
    schedule.every().day.at('21:51').do(greeting) # выполнять функцию greeting каждый день в определённое время
    schedule.every().thursday.do(greeting) # выполнять функцию greeting каждый вторник
    schedule.every().friday.at('23:45').do(greeting) # выполнять функцию greeting каждую пятницу в определённое время
    
    while True:
        schedule.run_pending()
    
    
if __name__ == '__main__':
    main()
