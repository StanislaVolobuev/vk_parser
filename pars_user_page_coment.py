import vk
import openpyxl
import time, datetime, os

id = '25053320'
 # по этому id осуществляется парсинг списка id из списка друзей


token = "token"
# token ключ доступа, необходимо
print('start')


head = [ 'id', 'first_name', 'political', 'people_main', 'last_seen', 'friends',
         'groups', 'videos', 'audios', 'photos', 'notes',
         'has_photo', 'coments'
]


def OpenFile(id):
    """
    На основании id определяет путь к файлу контента.
    :param id:
    :return:
    """
    print('open exel file')
    path = 'D:\\ira_data\\coment\\' + str(id) + '\\' +str(id) + '_user_coment.xlsx'
    wb_obj = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active
    row = 1
    data = []
    while row < sheet_obj.max_row+1:
        print(row, 'from', sheet_obj.max_row)
        data_row = []
        col = 1
        flag = True
        while col < sheet_obj.max_column:
            cell_obj = sheet_obj.cell(row=row, column=col)
            if cell_obj.value != None:
                data_row.append(cell_obj.value)

            col += 1
        row += 1

        data.append(data_row)

    return data
test_open = OpenFile(id)



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



def user_field(pars_id):
    """

    :param pars_id: id для парсинга данных страницы
    :return user_fields, born: данные со страницы, объект born
    нереализованная здесь наработка для определения года рождения методом поиска
    """
    time.sleep(1)
    vk_api = API(token)

    user_fields = vk_api.users.get(
        user_ids=pars_id,
        fields="deactivated, about, activities,  connections, counters, education, has_photo, is_favorite, is_hidden_from_feed, personal, trending,  last_seen"
    )

    return user_fields

def user_data_extract(row_pars):
    user_info = []
    for el in head:

        if el == 'coments':
            pass

        if el == 'id' or el == 'has_photo' or el == 'first_name':
            user_info.append(row_pars[0][el])

        if el == 'political' or el == 'people_main':
            if 'personal' in row_pars[0]:
                if len(row_pars[0]['personal']) > 0:
                    if el in row_pars[0]['personal']:
                        user_info.append(row_pars[0]['personal'][el])
                    else:
                        user_info.append('-')
                else:
                    user_info.append('-')
            else:
                user_info.append('-')


        if el == 'last_seen':
            if el in row_pars[0]:
                last_seen = row_pars[0]['last_seen']['time']
                user_info.append(str(datetime.date.fromtimestamp(last_seen)))
            else:
                user_info.append('-')

        if el == 'friends' or el == 'groups' or el == 'videos' or el == 'audios' or el == 'photos' or el == 'notes':
            if 'counters' in row_pars[0]:
                if len(row_pars[0]['counters']) > 0:

                    if el in row_pars[0]['counters']:
                        user_info.append(row_pars[0]['counters'][el])
                    else:
                        user_info.append('-')
                else:
                    user_info.append('-')
            else:
                user_info.append('-')

    return user_info


def start_work():
    data_post = OpenFile(id)
    list_ids = []
    data_result = []
    for row in data_post:
        if len(str(row[0])) > 1:
            row_result = []
            list_ids.append(row[0])
            user_pars = user_field(row[0])
            extract = user_data_extract(user_pars)
            for el in extract:
                row_result.append(el)
            for el in row[1:]:
                row_result.append(el)
            print(row_result, '-------------', extract)
            print(' ')
            data_result.append(row_result)

        else:
            pass



    print('write exel')
    wb = openpyxl.Workbook()

    # добавляем новый лист
    wb.create_sheet(title='Первый лист', index=0)
    # получаем лист, с которым будем работать
    sheet = wb['Первый лист']
    for col, title in zip(range(1, 20), head):
        cell = sheet.cell(row=1, column=col)
        cell.value = title
    for row, data_row in zip(range(2, len(data_result) + 2), data_result):
        for col, inform in zip(range(1, len(data_row) + 5), data_row):
            cell = sheet.cell(row=row, column=col)
            cell.value = inform

    wb.save('D:\\ira_data\\coment\\' + str(id) + '\\' +str(id) + '_page_user_coment.xlsx')

start = start_work()