from win32com.client import Dispatch
import datetime as dt
import time
# from check_registration import connecting_to_ecr

fr = Dispatch('Addin.DRvFR')

wait_cheque_timeout = 1 # время задержки для печати чека, может и не нужно если используем WaitForPrinting

print('Начало работы, подключение к ККТ')

# log_file_name = 'log_' + dt.datetime.isoformat(dt.datetime.now(), sep='_')[:-7] + '.txt'
log_file_name = 'log_' + str(dt.datetime.date(dt.datetime.now())) + '.txt'
logs_file_path = 'logs/' + log_file_name

# connecting_to_ecr()

def connecting_to_ecr():

    with open(logs_file_path, 'w+') as log:  # w - открытие (если нет - создается) файла на запись
        log.write(f'{dt.datetime.now()}: Начало тестирования ККТ \n')
        fr.GetECRStatus()
        if fr.ResultCode == 0:
            print('Подключение к ККТ прошло успешно')
            fr.TableNumber = 18
            fr.RowNumber = 1
            fr.FieldNumber = 1
            fr.ReadTable()
            log.write(
                f'{dt.datetime.now()}: Подключение к ККТ з\н {fr.ValueOfFieldString}, код ошибки: {fr.resultcode}, {fr.resultcodedescription}\n')
            print(f'ККТ з\н {fr.ValueOfFieldString}, код ошибки: {fr.resultcode}, {fr.resultcodedescription}')
            fr.GetDeviceMetrics()
            log.write(
                f'{dt.datetime.now()}: Модель ККТ {fr.UDescription}, прошивка {fr.ECRSoftVersion} от {dt.datetime.date(fr.ECRSoftDate)}\n')
            print(f'Модель ККТ {fr.UDescription}, прошивка {fr.ECRSoftVersion} от {dt.datetime.date(fr.ECRSoftDate)}')
            return True
        else:
            print(f'Подключение не удалось, код ошибки: {fr.resultcode}, {fr.resultcodedescription}')
            log.write(f'{dt.datetime.now()}: Подключение не удалось, код ошибки: {fr.resultcode}, {fr.resultcodedescription}\n')
            return False



def many_fn_operation_with_marking(number_of_positions=1, product_name='Товар', price=10, quantity=1):
    # пробитие чека с маркировкой
    print(f'Регистрируется кассовый чек с маркировкой с кол-вом позиций {number_of_positions}')

    with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
        log.seek(0, 2)
        fr.GetECRStatus()
        if fr.ECRMode == 2 or fr.ECRMode == 8:
            # fr.OpenCheck()
            if fr.resultcode != 0:
                print('After OpenCheck ', fr.resultcode, fr.resultcodedescription)
                fr.Disconnect()
                return

            for i in range(number_of_positions):
                fr.StringForPrinting = product_name
                fr.price = 1
                fr.quantity = 1
                fr.PaymentItemSign = 33
                fr.FNOperation()
                print(f'регистрация позиции {i + 1}, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                if fr.resultcode != 0:
                    print('After FNOperation ', fr.resultcode, fr.resultcodedescription)
                    fr.Disconnect()
                    return

                qr = "0102900021916404213Rfn-(uL4hLHv\x1D91EE06\x1D92ZL1qUSqxS/jylFxi1Sp/HouC05T7FqUi34uslMAoDc8="
                fr.BarCode = qr
                fr.ItemStatus = 1
                fr.FNSendItemBarcode()
                # time.sleep(1)
                print(f'Передача марки для позиции {i+1}, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                fr.FNGetDocumentSize()
                print(f'Размер текущего документа для ОФД, байт: {fr.DocumentSize}')
                print(f'Размер текущего уведомления о реализации маркированных товаров для ОИСМ, байт: {fr.NotificationSize}')
                log.write(
                    f'{dt.datetime.now()}: Передача марки для позиции {i+1}, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                log.write(f'Размер текущего документа для ОФД, байт: {fr.DocumentSize}\n')
                log.write(f'Размер текущего уведомления о реализации маркированных товаров для ОИСМ, байт: {fr.NotificationSize}\n\n')


                fr.TagNumber = 1262  # ИД. ФОИВ
                fr.TagType = 7
                fr.TagValueStr = "001"
                fr.FNSendTagOperation()
                fr.TagNumber = 1263  # ДАТА ДОК. ОСН.
                fr.TagType = 7
                fr.TagValueStr = "13.05.2024"
                fr.FNSendTagOperation()
                fr.TagNumber = 1264  # НОМЕР ДОК. ОСН.
                fr.TagType = 7
                fr.TagValueStr = "22"
                fr.FNSendTagOperation()
                fr.TagNumber = 1265  # ЗНАЧ. ОТР. РЕКВ.
                fr.TagType = 7
                fr.TagValueStr = "ЗНАЧ. ОТР. РЕКВ."
                fr.FNSendTagOperation()

            fr.Summ1 = 1000
            fr.CustomerEmail = 'buyer@mail.ru'
            fr.FNSendCustomerEmail()
            fr.PaymentTypeSign = 4  # ПризнакСпособаРасчета
            fr.FNCloseCheckEx()
            fr.WaitForPrinting()
            time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
            print(f'=============Закрытие чека==============\n{number_of_positions} позиций, код ошибки {fr.resultcode}, {fr.resultcodedescription}')
            log.write('=============Закрытие чека==============\n')
            log.write(
                f'{dt.datetime.now()}: Закрытие чека с маркировкой, {number_of_positions} позиций, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
            # result = self._get_cheque_from_fn()
            # log.write(f'Получен чек \n{result}')
            fr.Disconnect()
            return
        else:
            return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')

def many_fn_operation_minimal_check(number_of_positions=1, product_name='Товар', price=10, quantity=1):
    # пробитие чека с минимальными данными
    print(f'Регистрируется кассовый чек с минимальными данными с кол-вом позиций {number_of_positions}')

    with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
        log.seek(0, 2)
        fr.GetECRStatus()
        if fr.ECRMode == 2 or fr.ECRMode == 8:
            fr.OpenCheck()
            if fr.resultcode != 0:
                print('After OpenCheck ', fr.resultcode, fr.resultcodedescription)
                fr.Disconnect()
                return

            fr.CustomerEmail = 'buyer@mail.ru' # передаем email покупателя чтобы чек не печатался.
            fr.FNSendCustomerEmail()

            for i in range(number_of_positions):
                fr.StringForPrinting = product_name
                fr.price = 1
                fr.quantity = 1
                fr.PaymentItemSign = 1
                fr.FNOperation()
                print(f'регистрация позиции {i + 1}, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
                if fr.resultcode != 0:
                    print('After FNOperation ', fr.resultcode, fr.resultcodedescription)
                    fr.Disconnect()
                    return

            fr.Summ1 = 100000
            fr.PaymentTypeSign = 4  # ПризнакСпособаРасчета
            fr.FNCloseCheckEx()
            fr.WaitForPrinting()
            time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
            print(
                f'=============Закрытие чека==============\n{number_of_positions} позиций, код ошибки {fr.resultcode}, {fr.resultcodedescription}')
            log.write('=============Закрытие чека==============\n')
            log.write(
                f'{dt.datetime.now()}: Закрытие чека c минимальными данными, {number_of_positions} позиций, код ошибки {fr.resultcode}, {fr.resultcodedescription}\n')
            fr.Disconnect()
            return
        else:
            return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')


if __name__ == '__main__':

    connecting_to_ecr()
    # y = 0
    # while y != 'n':
    #     x = int(input('Введите кол-во позиций для чека - '))
    # many_fn_operation_with_marking(200)
    many_fn_operation_minimal_check(1000, 'т')
        # y = input('Еще разок ? y/n ').lower()

    fr.disconnect()