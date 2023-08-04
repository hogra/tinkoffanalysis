from datetime import timedelta, datetime
import xlsxwriter

from tinkoff.invest import CandleInterval, Client, exceptions
from tinkoff.invest.utils import now


TOKEN = input('Введите токен\n')
# TOKEN = 't.UFG_TMMUqB-NVhCLM4VU9umcsXv5CV7JaepKmxcpyETPzK0qmsWTU7cHCMIar__4n9cFydOZmyiRTOtCjwsSIQ'


def main():
    with Client(TOKEN) as client:
        print('Приступаю к работе')
        candles = list()
        print('Обрабатываю свечной график')
        for i in list(map(lambda x: x[6:-1],
                      list(filter(lambda x: 'figi' in x,
                                  str(client.instruments.get_favorites()).replace('(',
                                                                                  '|').replace(',', '|').split('|'))))):
            try:
                a = list(client.get_all_candles(figi=i, from_=now() - timedelta(days=365),
                                                interval=CandleInterval.CANDLE_INTERVAL_MONTH))
                candles.append({'name': client.instruments.share_by(id_type=1, id=i).instrument.name,
                                'year': a[0].open, 'hyear': a[6].open, 'pmonth': a[11].close,
                                'month': list(client.get_all_candles(figi=i,
                                                                     from_=now() - timedelta(minutes=1),
                                                                     interval=
                                                                     CandleInterval.CANDLE_INTERVAL_1_MIN))[-1].close})
            except exceptions.RequestError:
                pass
            except IndexError:
                pass
    print('Начинаю запись в xlsx')
    workbook = xlsxwriter.Workbook(f'{datetime.now().strftime("%d-%m-%Y %H-%M-%S")}.xlsx')
    worksheet = workbook.add_worksheet()
    data = [(i['name'],
             (((float(f"{i['month'].units}.{i['month'].nano}") - float(f"{i['pmonth'].units}.{i['pmonth'].nano}")) /
               float(f"{i['pmonth'].units}.{i['pmonth'].nano}")) * 100),
             (((float(f"{i['month'].units}.{i['month'].nano}") - float(f"{i['hyear'].units}.{i['hyear'].nano}")) /
               float(f"{i['hyear'].units}.{i['hyear'].nano}")) * 100),
             (((float(f"{i['month'].units}.{i['month'].nano}") - float(f"{i['year'].units}.{i['year'].nano}")) /
               float(f"{i['year'].units}.{i['year'].nano}")) * 100)) for i in candles]
    worksheet.write(0, 0, 'Компания')
    worksheet.write(0, 1, 'Месяц')
    worksheet.write(0, 2, 'Полгода')
    worksheet.write(0, 3, 'Год')
    for row, (item, month, hyear, year) in enumerate(data):
        worksheet.write(row + 1, 0, item)
        worksheet.write(row + 1, 1, month)
        worksheet.write(row + 1, 2, hyear)
        worksheet.write(row + 1, 3, year)
        row += 1
    workbook.close()
    print('Закончил работу')
    input()
    return 0


if __name__ == "__main__":
     main()