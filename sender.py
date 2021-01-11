# -- coding: utf-8 --
import from_excel
import smtplib
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart  # Многокомпонентный объект
from email.mime.text import MIMEText  # Текст/HTML

# *** - звездочки заменить не убирая ковычки

letter = "***"  # Текст письма для рассылки


def start_sending(mail_list):
    """ Функция производит рассылку шаблонного письма вставляя имя и отчество адресата. Для поддтверждения отправки
    скрытая копия высылается на технический почтовый ящик. В качестве аргумента функция принимает список со вложенными
    списками, содержащими строки ФИО и адрес эл. почты. При нескольких адресах эл. почты у одного получателя эти адреса
    должны быть вложены в список. Если фамилии нет, то нужно добавить любое слово перед именем отчеством.
    Пример аргумента:
    [['Фамилия Имя отчество', 'email@domen.ru'], ['Фамилия Имя Отчество', ['email1@domen.ru', 'email2@omen.ru']].
     """
    for company in mail_list:
        today = int(datetime.strftime(datetime.now(), "%w"))
        while today == 0 or today == 6:
            time.sleep(3600)
            today = int(datetime.strftime(datetime.now(), "%w"))
        current_hour = int(datetime.strftime(datetime.now(), "%H"))
        while current_hour < 8 or current_hour > 18:
            time.sleep(3600)
            current_hour = int(datetime.strftime(datetime.now(), "%H"))
        print(from_excel.client_base.index(company))
        if type(company[1]) == list:
            for i in range(len(company[1])):
                address_from = "***@***.**"  # Почта, с которой ведется рассылка
                address_to = company[1][i]
                address_bcc = '***@***.**'  # Скрытая копия на технический ящик
                password = "***"  # Пароль от почты, с которой ведется рассылка
                msg = MIMEMultipart()
                msg['From'] = address_from
                msg['To'] = address_to
                msg['Subject'] = '***'  # Тема письма
                msg['Bcc'] = address_bcc
                name = company[0].split(sep=' ')
                name = ' '.join(name[1:])
                body = "Добрый день, {}! \n".format(name) + letter
                msg.attach(MIMEText(body, 'plain'))
                server = smtplib.SMTP_SSL('smtp.**.ru', 465)  # Почтовый сервер и порт
                # server.set_debuglevel(True)  # Режим отладки
                # server.starttls()  # Шифрованный обмен по TLS необходим для некоторых почтовых серверов
                server.login(address_from, password)
                server.send_message(msg)
                server.quit()
                time.sleep(1800)  # Пауза между отправкой писем в секундах
                # По субботам и воскресеньям пиьсма не отправляются
                today = int(datetime.strftime(datetime.now(), "%w"))
                while today == 0 or today == 6:
                    time.sleep(3600)
                    today = int(datetime.strftime(datetime.now(), "%w"))
                current_hour = int(datetime.strftime(datetime.now(), "%H"))
                while current_hour < 8 or current_hour > 18:
                    time.sleep(3600)
                    current_hour = int(datetime.strftime(datetime.now(), "%H"))
        else:
            address_from = "***@**.**"    # Почта, с которой ведется рассылка
            address_to = company[1]
            address_bcc = '***@**.**'  # Скрытая копия на технический ящик
            password = "***"  # Пароль от почты, с которой ведется рассылка
            msg = MIMEMultipart()
            msg['From'] = address_from
            msg['To'] = address_to
            msg['Subject'] = '***'  # Тема письма
            msg['Bcc'] = address_bcc
            name = company[0].split(sep=' ')
            name = ' '.join(name[1:])
            body = "Добрый день, {}! \n".format(name) + letter
            msg.attach(MIMEText(body, 'plain'))
            server = smtplib.SMTP_SSL('smtp.**.ru', 465)  # Почтовый сервер и порт
            # server.set_debuglevel(True)  # Режим отладки
            # server.starttls()  # Шифрованный обмен по TLS необходим для некоторых почтовых серверов
            server.login(address_from, password)
            server.send_message(msg)
            server.quit()
            time.sleep(1800)  # Пауза между отправкой писем в секундах


start_sending(from_excel.client_base)
