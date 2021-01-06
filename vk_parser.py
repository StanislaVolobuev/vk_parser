import vk
import openpyxl

my_id = 413569747
def new_session():
    '''
    :param my_id: Принимает id пользователя, от имени
    которого будут осуществляться запросы к VK API
    :return: обьект сессии vk_api
    '''
    session = vk.Session(access_token="a58d2d18787a6943e347cc0f852daaef276e2e08a8ec5f17f7712acf09dc88f7b5f225b18c93f88cf4c4f")
    vk_api = vk.API(
        session,  v = '5.35' ,
        lang = 'ru' ,
        timeout = 10
                    )
    return vk_api


def user_information(vk_api, use_id=my_id):
    '''

    :param vk_api: открытая сессия для запросов вконтакте
    :param use_id: принимаемый id пользователя, для
    сбора данных
    :return: user_fields: словарь полей с данными о пользоватиле
    '''
    user_fields = vk_api.users.get(
        user_id=use_id,
        fields="deactivated, is_closed, can_access_closed, about, activities, bdate, books, can_post, can_see_audio, can_send_friend_request, can_write_private_message, career, city, connections, contacts, counters, country, domain, education, exports, followers_count, friend_status, games, has_mobile, has_photo, home_town, interests, is_favorite, is_friend, is_hidden_from_feed, last_seen, lists,maiden_name, military, movies, music, nickname, occupation, personal, quotes, relatives, relation, schools, screen_name, site,   sex, status, timezone, trending, tv, universities, verified, wall_default"
        )
    print(user_fields)
    return user_fields


def friends_list(vk_api, use_id=my_id):
    user_friends = vk_api.friends.get(
        user_id=use_id,
        order='hints'
    )
    print(user_friends)
    return user_friends





def main():
    pass

if __name__ == '__main__':
    main()




