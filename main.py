import openpyxl
import random
import re
import smtplib
from email.mime.text import MIMEText
from email.header import Header

# start_line = 1
max_mail = int(input('Сколько адресов нужно: '))


def create_sender_list():
    start_line = 1
    wb = openpyxl.Workbook()
    column_a = 'A'
    column_b = 'B'
    while start_line <= max_mail:
        sheets_list = wb.sheetnames  # Получаем список всех листов в файле
        sheet_active = wb[sheets_list[0]]  # Начинаем работать с самым первым
        mail_server_list = ['gmail.com', 'yandex.ru', 'outlook.com', 'mail.ru']
        random_value = random.randrange(1, 10)  # Генерируем случайное число от 1 до 10
        random_mail = random.sample('abcdefghijklmnopqrstuvwxyz0123456789',
                                    random_value)  # Генерируем случайный адрес из набора символов
        random_mail = ''.join(random_mail)
        random_mail_server = mail_server_list[
            random.randrange(0, len(mail_server_list))]  # Выбираем случайный почтовый сервер
        random_mail = random_mail + '@' + random_mail_server  # Создаем итоговый адрес
        start_line = start_line + 1
        start_line = str(start_line)
        sheet_active[column_a + start_line] = random_mail
        start_line = int(start_line)

        # Генерируем пароли к почте
        random_value = random.randrange(8, 12)  # Генерируем длину пароля (от 8 до 12 символов)
        password_for_mail = random.sample('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*',
                                          random_value)  # Генерируем сам пароль с использованием символов в скобках
        password_for_mail = ''.join(password_for_mail)

        # А теперь пишем все в файл
        start_line = str(start_line)
        sheet_active[column_b + start_line] = password_for_mail
        start_line = int(start_line)
        print(random_mail, ':', password_for_mail, ' - создан')

        wb.save('sender_base.xlsx')
    print('База почтовых адресов для отправки создана.\n')
