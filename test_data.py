import vk

id = 11132894
token = "token"

def API(id, token):
    session = vk.Session(access_token=token)
    vk_api = vk.API(
        session,  v = '5.35' ,
        lang = 'ru' ,
        timeout = 10
    )
    return vk_api


def user_field():
    vk_api = API(id, token)
    user_fields = vk_api.users.get(
        user_id=id,
        fields="deactivated, is_closed, can_access_closed, about, activities, bdate, books, can_post, can_see_audio, can_send_friend_request, can_write_private_message, career, city, connections, contacts, counters, country, domain, education, exports, followers_count, friend_status, games, has_mobile, has_photo, home_town, interests, is_favorite, is_friend, is_hidden_from_feed, lists,maiden_name, military, movies, music, nickname, occupation, personal, quotes, relatives, relation, schools, screen_name, site,   sex, status, timezone, trending, tv, universities, verified, wall_default"
    )


    user_wall_own = vk_api.wall.get(
        user_id=11132894, filter = 'owner', count = "1")

    user_wall_all = vk_api.wall.get(
        user_id=11132894, filter = 'all', count = "1")
    user_fields[0]['wall_own'] = user_wall_own['count']
    user_fields[0]['wall_all'] = user_wall_all['count']
    print(user_fields)
    return user_fields
p = user_field()

def data_change(p):
    '''
    функция подготовки данных для переноса в эксель
    используя список head (его значения - названия столбцов) находит соответствующую
    информацию в массиве и формирует соответствующий список для строчного переноса в эксель
    :param p:
    :return:head - список полей для обработки данных и внесения в эксель
            data - список содержащий данные в соответствии со списком head
    '''
    head = [
        'first_name', 'id', 'last_name', 'sex', 'screen_name', 'verified',
        'friend_status', 'nickname', 'maiden_name', 'domain', 'bdate',
        'city', 'country', 'timezone', 'has_photo', 'has_mobile',
        'is_friend', 'can_post', 'can_see_audio', 'skype', 'wall_default',
        'interests', 'books', 'tv', 'quotes', 'about', 'games', 'movies',
        'activities', 'music', 'can_write_private_message',
        'can_send_friend_request', 'mobile_phone', 'home_phone', 'site',
        'status', 'exports', 'followers_count', 'is_favorite',
        'is_hidden_from_feed', 'occupation', 'career', 'military',
        'university', 'university_name', 'faculty', 'faculty_name',
        'graduation', 'education_form', 'education_status', 'home_town',
        'relation', 'personal', 'alcohol', 'inspired_by', 'langs',
        'life_main', 'people_main', 'religion', 'smoking', 'universities',
        'schools', 'relatives', 'counters', 'albums', 'audios', 'followers',
        'friends', 'gifts', 'notes', 'online_friends', 'pages', 'photos',
        'subscriptions', 'user_photos', 'videos', 'new_photo_tags',
        'new_recognition_tags', 'clips_followers', 'groups', 'wall_own', 'wall_all'
    ]
    data = list()
    header = list()
    print('gjhhkjjlijkjgghc', len(head), head)
    num = 0
    while num < 83:
        if num < 11 or 12 < num < 36 or 36 < num <40 or 42 < num < 52 or 79 < num < 82:
            data.append(p[0][head[num]])
        if num == 11 or num == 12:
            data.append(p[0][head[num]]['title'])
        if num == 36:
            if len(p[0][head[num]]) > 0:
                data.append(' '.join(p[0][head[num]]))
        if num == 40:
            data_str = ''
            if len(p[0][head[num]]) > 0:
                for key in p[0][head[num]].keys():
                    data_str += ' '
                    data_str += str(p[0][head[num]][key])
            data.append(data_str)
        if num == 41 or num == 42 or num == 60 or num == 61 or num == 62:
            data_str = ''
            if len(p[0][head[num]]) > 0:
                # test = []
                for i in p[0][head[num]]:
                    # test.append(i)
                    data_dict = i
                    dda = ' '
                    for key in data_dict.keys():
                        dda += ' '
                        dda += str(data_dict[key])
                    print(dda)

                    '''
                    for key in p[0][head[num]][i].keys():
                        data_str += ' '
                        data_str += str(p[0][head[num]][i][key])
                   
                    '''
                    data_str += dda
                    data_str += '; '

                data.append(data_str)
        if num == 52:
            if len(p[0][head[num]]) > 0:
                data.append('=>')
                while num != 59:
                    num += 1
                    data_personal = p[0][head[52]]
                    if num == 55:
                        langs_inf = ' '.join(data_personal[head[num]])
                        data.append(langs_inf)
                    else:
                        data.append(data_personal[head[num]])
        if num == 63:
            data_counters = p[0][head[num]]
            data.append('=>')
        if 63 < num < 80:
            data_counters == p[0][head[63]]
            data.append(data_counters[head[num]])

        num += 1
    print('data', data)
    return data, head


test2 = data_change(p)


