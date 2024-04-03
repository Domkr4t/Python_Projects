from forex_python.bitcoin import BtcConverter
from forex_python.converter import CurrencyRates

api = CurrencyRates()
btc = BtcConverter()

btc_price = btc.get_latest_price('USD')
dollar_rate = api.get_rate('RUB', 'USD')
convert = api.convert ('RUB', 'USD', 1000)

print(btc_price, dollar_rate, convert)
