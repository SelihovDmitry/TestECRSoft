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


def write_fonts_in_table(font=1):
    #заполняем таблицу шрифтов указанным шрифтом

    for field in range(1, 26):
        fr.TableNumber = 8
        fr.RowNumber = 1
        fr.FieldNumber = field
        fr.ValueOfFieldInteger = font
        fr.WriteTable()
        if fr.resultcode != 0:
            print('After WriteTable ', fr.resultcode, fr.resultcodedescription)
            fr.Disconnect()
            return


def check_different_fonts_and_pattern(number_of_positions=1, product_name='Товар', current_font=1, price=10, quantity=1):
    # пробитие чеков с разными шаблонами и шрифтами

    for compact_header in range(10): # перебираем компактные заголовки от 0 до 9
        # записываем значение компактного заголовка в Т17П12
        fr.TableNumber = 17
        fr.RowNumber = 1
        fr.FieldNumber = 18
        fr.ValueOfFieldInteger = compact_header
        fr.WriteTable()
        if fr.resultcode != 0:
            print('After WriteTable ', fr.resultcode, fr.resultcodedescription)
            fr.Disconnect()
            return

        for pattern_ending in range(10):
            if pattern_ending != 7:
                # записываем значение шаблона окончания в Т17П56
                fr.TableNumber = 17
                fr.RowNumber = 1
                fr.FieldNumber = 56
                fr.ValueOfFieldInteger = pattern_ending
                fr.WriteTable()
                if fr.resultcode != 0:
                    print('After WriteTable ', fr.resultcode, fr.resultcodedescription)
                    fr.Disconnect()
                    return

                print(f'Пробиваем чек шрифтом {current_font} с компактным заголовком {compact_header} и шаблоном окончания {pattern_ending}')
                fr.StringForPrinting = '*************************'
                fr.PrintString()
                fr.StringForPrinting = f'ПРОБИВАЕМ ЧЕК ШРИФТОМ ------ {current_font}'
                fr.PrintString()
                fr.StringForPrinting = f'с компактным заголовком ---- {compact_header}'
                fr.PrintString()
                fr.StringForPrinting = f'и шаблоном окончания ------- {pattern_ending}'
                fr.PrintString()
                fr.StringForPrinting = '*************************'
                fr.PrintString()

                fr.GetECRStatus()
                if fr.ECRMode == 2:
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
                        if fr.resultcode != 0:
                            print('After FNOperation ', fr.resultcode, fr.resultcodedescription)
                            fr.Disconnect()
                            return

                    fr.Summ1 = 100
                    fr.PaymentTypeSign = 4  # ПризнакСпособаРасчета
                    fr.StringForPrinting = ''
                    fr.FNCloseCheckEx()
                    print(f'====Закрытие чека с компактным заголовком {compact_header} ====\n{number_of_positions} позиций, '
                          f'код ошибки {fr.resultcode}, {fr.resultcodedescription}')
                    if fr.resultcode != 0:
                        print('After FNCloseCheckEx ', fr.resultcode, fr.resultcodedescription)
                        fr.CancelCheck()
                    fr.WaitForPrinting()
                    # time.sleep(wait_cheque_timeout)  # задержка - даем время на печать на всякий случай
                else:
                    return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')
    fr.Disconnect()

def check_with_new_tax(Tax):
    if fr.ECRMode == 2 or fr.ECRMode == 8:
        fr.StringForPrinting = 'Наименование товара'
        fr.price = 100
        fr.quantity = 1
        fr.tax1 = 1
        fr.FNOperation()
        print('After FNOperation ', fr.resultcode, fr.resultcodedescription)

        fr.Summ1 = 1000
        fr.CustomerEmail = 'buyer@mail.ru'
        fr.FNSendCustomerEmail()
        fr.FNCloseCheckEx()
        print('After FNCloseCheckEx ', fr.resultcode, fr.resultcodedescription)
        fr.WaitForPrinting()
        fr.Disconnect()
        return
    else:
        return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')


def main():
    fonts = [1]
    for font in fonts:
        write_fonts_in_table(font)
        check_different_fonts_and_pattern(number_of_positions=1, product_name='Наименование товара', current_font=font)


if __name__ == '__main__':
    # for compact_header in range(10):
    #     if compact_header != 7:
    #         print(compact_header)

    connecting_to_ecr()
    # check_with_new_tax(1)
    main()
    # check_different_fonts_and_pattern(1, 'Наименование товара')
    # write_fonts_in_table(1)



