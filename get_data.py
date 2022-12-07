import requests
import json

# data = requests.get('https://kaspi.kz/mc/api/orderTabs/active?count=50&selectedTabs=DELIVERY&startIndex=0&returnedToWarehouse=false')
# print(data)


async def make_request():
    s = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36",
    }
    data = {
        "action": "login",
        "username": "alevtinanur89@gmail.com",
        "password": "Dmitrii8989!)"
    }
    login_url = 'https://kaspi.kz/mc/api/login'

    s.post(
        url=login_url,
        headers=headers,
        data=data
    )
    return s


async def get_delivery_data_from_api():
    try:
        s = await make_request()
        headers2 = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36",
            'Content-Type': 'application/json',
        }
        delivery_url = 'https://kaspi.kz/mc/api/orderTabs/active?count=2000&selectedTabs=DELIVERY&startIndex=0&returnedToWarehouse=false'
        detail = 'https://kaspi.kz/merchantcabinet/api/order/details/%7Bstatus%7D/229325797'

        request_data = s.get(
            delivery_url,
            headers=headers2,
            cookies=s.cookies.get_dict()
        )

        with open('delivery.json', 'w', encoding='utf-8') as my_json:
            json.dump(request_data.json(), my_json, ensure_ascii=False, indent=4)

        with open('delivery.json') as json_file:
            delivery_data = json.load(json_file)

        for prod in delivery_data[0]["orders"]:
            detail_data = s.get(
                'https://kaspi.kz/merchantcabinet/api/order/details/%7Bstatus%7D/' + str(prod["orderCode"]),
                headers=headers2,
                cookies=s.cookies.get_dict()
            )
            with open(f'{prod["orderCode"]}_prod_detail.json', 'w', encoding='utf-8') as my_json:
                json.dump(detail_data.json(), my_json, ensure_ascii=False, indent=4)
    except KeyError:
        return 'empty'


async def get_pickup_data_from_api():
    try:
        s = await make_request()
        headers2 = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36",
            'Content-Type': 'application/json',
        }
        pickup_url = 'https://kaspi.kz/mc/api/orderTabs/active?count=50&selectedTabs=PICKUP&startIndex=0&returnedToWarehouse=false'
        detail = 'https://kaspi.kz/merchantcabinet/api/order/details/%7Bstatus%7D/229325797'

        request_data = s.get(
            pickup_url,
            headers=headers2,
            cookies=s.cookies.get_dict()
        )

        with open('pickup.json', 'w', encoding='utf-8') as my_json:
            json.dump(request_data.json(), my_json, ensure_ascii=False, indent=4)

        with open('pickup.json') as json_file:
            pickup_data = json.load(json_file)


        for prod in pickup_data[0]["orders"]:
            detail_data = s.get(
                'https://kaspi.kz/merchantcabinet/api/order/details/%7Bstatus%7D/' + str(prod["orderCode"]),
                headers=headers2,
                cookies=s.cookies.get_dict()
            )
            with open(f'{prod["orderCode"]}_pickup_prod_detail.json', 'w', encoding='utf-8') as my_json:
                json.dump(detail_data.json(), my_json, ensure_ascii=False, indent=4)
    except KeyError:
        return 'empty'

