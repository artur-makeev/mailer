import pandas as pd


def load_from_excel(file):
    """ Функция формирует список из списков, состоящих из фамилии имени отчества и адреса электронной почты. Если
    адресов электронной почты у одного контакта несколько, то формируется список из адресов эл. почты.
    Пример вывода:
    [['Фамилия Имя отчество', 'email@domen.ru'], ['Фамилия Имя Отчество', ['email1@domen.ru', 'email2@omen.ru']].
    В качестве аргумента функци принимается файл эксель.
    """
    xls = pd.ExcelFile(file)
    df = xls.parse(0, usecols=[2, 6])  # usecols=[колонка с ФИО, колонка с адресами эл. почты]
    contacts_draft = [[df.loc[n][0], df.loc[n][1]] for n in range(0, len(df.index))]
    contacts_final = []
    for elem in contacts_draft:  # Преобразуем строки с несколькими эл. почтами, разделенные символом \n в списки
        temp_list = [elem[0]]
        if '\n' in elem[1]:
            temp_list.append(elem[1].split('\n'))
            contacts_final.append(temp_list)
        else:
            temp_list.append(elem[1])
            contacts_final.append(temp_list)
    return contacts_final


client_base = load_from_excel('segment.xlsx')
