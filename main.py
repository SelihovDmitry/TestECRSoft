from check_registration import ECR
from check_registration import connecting_to_ecr
from check_registration import logs_file_path
from fixed_data import kkt_tags
import datetime as dt

def check_tags(document, required_tags):
    # функция проверки вхождения тега в ЭФ чека
    current_doc = document.split('\n')[0]
    print(f'Проверка тегов в документе {current_doc}')
    with open(logs_file_path, 'r+') as log:  # r+ - открытие файла на чтение и изменение
        log.seek(0, 2)
        with open('ERROR.txt', 'w+') as errlog:
            # errlog.write(dt.datetime.now())
            for tag in required_tags:
                if str(tag) in document:
                    log.write(f'Тег {tag} есть в чеке\n')
                else:
                    log.write(f'ERROR!!! Тега {tag} нет в чеке !!!\n')
                    errlog.write(f'{dt.datetime.now()} : ERROR!!! Тега {tag} нет в чеке {current_doc}!!!\n')

def main():

    ECROnTest = ECR()

    connecting_to_ecr()

    registration_report = ECROnTest.registration_report()
    check_tags(registration_report, kkt_tags.reg_tags)

    open_session_cheque = ECROnTest.open_session()
    check_tags(open_session_cheque, kkt_tags.open_session_tags)

    simple_cheque = ECROnTest.fn_operation_min()
    check_tags(simple_cheque, kkt_tags.min_cheque_tags)

    cheque_with_marking = ECROnTest.fn_operation_with_marking()
    check_tags(cheque_with_marking, kkt_tags.cheque_with_marking)

    close_session_cheque = ECROnTest.close_session()
    check_tags(close_session_cheque, kkt_tags.close_session_tags)

    calculation_state_report = ECROnTest.calculation_state_report()
    check_tags(calculation_state_report, kkt_tags.otchet_o_sost_rasch_tags)


if __name__ == '__main__':
    print('Hello you in module main')
    main()