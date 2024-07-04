# Программа тестирования прошивок ККТ

TestECRSoft - программа тестирования прошивок ККТ на протоколе Штрих-М.

Язык программирования - Python 3.12\
IDE - PyCharm Community Edition 2024.1.3

Модуль main.py - основной модуль.\
Функция check_tags - на вход подается чек из ФН в виде текста и список тегов в виде списка

Модуль check_registration.py - в нем реализовано подключение драйвера ККТ и класс
ECR в котором методы регистрации документов (чеки, отчеты).
Методы регистрируют документ и возвращают его в электронной форме полученной из ФН после регистрации.

Реализовано ведение лога (в папку logs) и отдельно лога ошибок 

Модуль kkt_tags из пакета fixed_data - теги ККТ по типам чеков.

## Как использовать

- подключить ФР к ПК и настроить с ним связь из Теста драйвера
- ФР должен быть фискализирован и находиться в режиме 4 (закрытая смена)
- в каталоге с программой надо создать папку logs
- перед запуском модуля надо закрыть Тест драйвера или на вкладке Прочее - Связь - Нажать разорвать связь
