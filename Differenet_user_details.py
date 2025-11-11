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

def check_with_diff_user_details(mesto_rasch='Место расчетов по умолчанию', addr_rasch='Адрес расчетов по умолчанию', kassir='Имя Кассира По Умолчанию'):

    print(
        f'Пробиваем чек c длиной места расч. {len(mesto_rasch)}, длиной адр. расч.{len(addr_rasch)} длина имя кассира {len(kassir)}')
    fr.StringForPrinting = '*************************'
    fr.PrintString()
    fr.StringForPrinting = f'ПРОБИВАЕМ ЧЕК c длиной места расч. {len(mesto_rasch)}'
    fr.PrintString()
    fr.StringForPrinting = f'длиной адреса расчетов {len(addr_rasch)}'
    fr.PrintString()
    fr.StringForPrinting = f'и длина имя кассира {len(kassir)}'
    fr.PrintString()
    fr.StringForPrinting = '*************************'
    fr.PrintString()

    if fr.ECRMode == 2 or fr.ECRMode == 8:
        fr.StringForPrinting = 'Наименование товара'
        fr.price = 1.11
        fr.quantity = 1
        fr.tax1 = 1
        fr.FNOperation()
        print('After FNOperation ', fr.resultcode, fr.resultcodedescription)

        fr.TagNumber = 1187  # Место расчетов
        fr.TagType = 7
        fr.TagValueStr = mesto_rasch
        fr.FNSendTag()

        fr.TagNumber = 1009  # Адрес расчетов
        fr.TagType = 7
        fr.TagValueStr = addr_rasch
        fr.FNSendTag()

        fr.TagNumber = 1021  # Имя кассира
        fr.TagType = 7
        fr.TagValueStr = kassir
        fr.FNSendTag()

        print('After FNSendTagOperation ', fr.resultcode, fr.resultcodedescription)

        fr.Summ1 = 1000
        fr.CustomerEmail = 'buyer@mail.ru'
        fr.FNSendCustomerEmail()
        fr.FNCloseCheckEx()
        print('After FNCloseCheckEx ', fr.resultcode, fr.resultcodedescription)
        if fr.resultcode != 0:
            print('After FNCloseCheckEx ', fr.resultcode, fr.resultcodedescription)
            fr.CancelCheck()
        fr.WaitForPrinting()
        # fr.Disconnect()
        return
    else:
        return print(f'ККТ не в режиме 2, режим ККТ: {fr.ECRMode}')

def main():

    diff_mest_rasch = ['M','Средней длины место расчетов с покупателем',
                       'Самоеееееее длиииннное значееение адреса !!!! МЕСТА !!!! расчетов 128 символов 0123456789012345678901234567890123456789012_МЕСТО']
    diff_addr_rasch = ['A','Средней длины адрес расчетов с покупателем',
                       'Самоееееее длиииннное значееение адреса !!!! АДРЕСА !!!! расчетов 128 символов 0123456789012345678901234567890123456789012_АДРЕС']
    diff_kassir = ['K','Среднее Имя Кассира','Длинное Имя Кассира 64 символа Красильникова Капитолина Конст.64']

    # for k in diff_mest_rasch:
    #     print(len(k))
    # for k in diff_addr_rasch:
    #     print(len(k))
    # for k in diff_kassir:
    #     print(len(k))

    for mest_rasch in diff_mest_rasch:
        for addr_rasch in diff_addr_rasch:
            for kassir in diff_kassir:
                check_with_diff_user_details(mest_rasch, addr_rasch, kassir)



    # check_with_diff_user_details()

if __name__ == '__main__':
    connecting_to_ecr()
    main()
    fr.Disconnect()
