import vk
import openpyxl
import time, datetime, os

id = 'novostinl' # по этому id осуществляется парсинг списка id из списка друзей
token = "token"
# token ключ доступа, необходимо



def OpenFile(id):
    """
    На основании id определяет пут к файлу контента.
    :param id:
    :return:
    """

    path = 'data\\'+id+'\\'+id+'_content.xlsx'
    wb_obj = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active
    row = 1
    data = []
    while row < sheet_obj.max_row+1:
        data_row = []
        col = 1
        if row == 1:
            while col < 11:
                cell_obj = sheet_obj.cell(row=row, column=col)
                data_row.append(cell_obj.value)
                col += 1
            data.append(data_row)

        else:
            cell_coment = sheet_obj.cell(row=row, column=6)
            if int(cell_coment.value) > 3:

                while col < 11:
                    cell_obj = sheet_obj.cell(row=row, column=col)
                    data_row.append(cell_obj.value)
                    col += 1
            if len(data_row) > 0:
                data.append(data_row)
        row += 1
    return  data


def API(token):
    """

    :param token: ключ с правами доступа
    :return vk_api: открытая сессия для запросов к серверу
    """
    session = vk.Session(access_token=token)
    vk_api = vk.API(
        session,  v = '5.126' ,
        lang = 'ru' ,
        timeout = 100
    )
    return vk_api


start = OpenFile(id)
num = 1
for el in start:
    print(num, el)
    num += 1