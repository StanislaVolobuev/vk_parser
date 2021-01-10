import  vk
import openpyxl


my_id = 413569747
def user_f(id):
    session = vk.Session(access_token="token")
    vk_api = vk.API(
        session,  v = '5.35' ,
        lang = 'ru' ,
        timeout = 10
    )
    user_fields = vk_api.users.get(
        user_id=id,
        fields="deactivated, is_closed, can_access_closed, about, activities, bdate, books, can_post, can_see_audio, can_send_friend_request, can_write_private_message, career, city, connections, contacts, counters, country, domain, education, exports, followers_count, friend_status, games, has_mobile, has_photo, home_town, interests, is_favorite, is_friend, is_hidden_from_feed, lists,maiden_name, military, movies, music, nickname, occupation, personal, quotes, relatives, relation, schools, screen_name, site,   sex, status, timezone, trending, tv, universities, verified, wall_default"
    )
    print(user_fields)
    return user_fields


p = user_f(413569747)
list_keys = list(p[0].keys())
print(list_keys, len(list_keys))
for key in list(p[0].keys()):
    print(p[0][key])


def data_change(p):
    data = list()
    header = list()
    for key in list(p[0].keys()):
        if key  in ['city', 'country']:
            print('ok', p[0][key]['title'])
            # data[key] = p[0][key]['title']
            header.append(key)
            data.append(p[0][key]['title'])
        elif key  in ['career', 'military', 'personal', 'universities', 'schools', 'relatives']:
            print('ok', ' '.join(p[0][key]))
            # data[key] = ' '.join(p[0][key])
            header.append(key)
            data.append(' '.join(p[0][key]))
        elif key == 'counters':
            for i in list(p[0][key].keys()):
                print('ok', p[0][key][i])
                # data[i] = p[0][key][i]
                header.append(i)
                data.append(p[0][key][i])
        else:
            print('ok', p[0][key])
            header.append(key)
            data.append(p[0][key])
    print(header)
    print(data)
        # data[key] = p[0][key]



data_use =data_change(p)

def friends_list(use_id=my_id):
    session = vk.Session(access_token="a58d2d18787a6943e347cc0f852daaef276e2e08a8ec5f17f7712acf09dc88f7b5f225b18c93f88cf4c4f")
    vk_api = vk.API(
        session,  v = '5.35' ,
        lang = 'ru' ,
        timeout = 10
    )
    user_friends = vk_api.friends.get(
        user_id=use_id,
        order='hints'
    )
    print(user_friends)
    return user_friends


def matrix():
    data_user = user_f(413569747)
    list_keys = list(p[0].keys())
    # создаем новый excel-файл
    wb = openpyxl.Workbook()

    # добавляем новый лист
    wb.create_sheet(title='Первый лист', index=0)
    # получаем лист, с которым будем работать
    sheet = wb['Первый лист']

    for col, key in zip(range(1, len(list_keys)+2), list_keys):
        cell = sheet.cell(row=1, column=col)
        cell.value = key
        if key == 'city':
            p[0][key] = p[0][key]['title']
        # cell = sheet.cell(row=2, column=col)
        # cell.value = p[0][key]
        print(key, p[0][key])
    '''
    list_friends = friends_list()
    n=0
    for row, friend in zip(range(2, len(list_friends['items']) + 2),list_friends['items']):
        data_user = user_f(friend)
        n +=1
        for col, key in zip(range(1, len(list_keys) + 3), list_keys):
            cell = sheet.cell(row=row, column=col)
            cell.value = data_user[0][key]
            if n > 10:
                break
    
    '''
    wb.save('example.xlsx')



