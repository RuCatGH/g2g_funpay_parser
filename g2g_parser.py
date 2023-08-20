import aiohttp
import asyncio
import jmespath
from openpyxl import Workbook
import sys

wb = Workbook()
ws = wb.active
ws.append(['1st tier category', '2nd tier category', 'offernum 2nd tier', 'filters 2nd tier', '3rd tier category', 'offersnum 3rd tier'])


async def make_request(url, params=None):
    async with aiohttp.ClientSession() as session:
        async with session.get(url, params=params) as response:
            response.raise_for_status()
            return await response.json()

# Парсинг фильтров
async def get_filters(service_id, brand, keywords_json):
    params = {
        'service_id': service_id,
        'brand_id': brand,
    }
    # Парсинг фильтров
    try:
        response = await make_request('https://sls.g2g.com/offer/keyword_relation/collection', params=params)
    except aiohttp.ClientError:
        await asyncio.sleep(1)
        response = await make_request('https://sls.g2g.com/offer/keyword_relation/collection', params=params)
        
    relation_collection = response

    filter_2nd = {}
    if jmespath.search('payload.results[]', relation_collection):
        for result in relation_collection['payload']['results']:
            label = result['label']['en']
            if result['is_multi_layer']:
                name_group_filters = []
                for child in result['children']:
                    name_group_filter = child['value']
                    name_filters = []
                    for child_2 in child['children']:
                        name_filter = child_2['value'].replace("\xa0", " ")
                        name_filters.append(name_filter)
                    name_group_filters.append({name_group_filter: name_filters})
                filter_2nd[label] = name_group_filters
            else:
                name_group_filter = label
                name_filters = []
                for child in result['children']:
                    name_filter = child['value'].replace("\xa0", " ")
                    name_filters.append(name_filter)
                filter_2nd[label] = {name_group_filter: name_filters}

    params = {
        'service_id': service_id,
        'brand_id': brand,
        'country': 'RU',
    }
    # Если есть фильтров регионов, также добавляем
    try:
        region_response = await make_request('https://sls.g2g.com/offer/keyword_relation/region', params=params)
    except aiohttp.ClientError:
        await asyncio.sleep(1)
        region_response = await make_request('https://sls.g2g.com/offer/keyword_relation/region', params=params)
        
    regions = jmespath.search('payload.results[].region_id', region_response)

    if regions:
        regions_filter = [keywords_json[region]['en'] for region in regions]
        filter_2nd['Region'] = regions_filter

    return filter_2nd

# Основная функция
async def get_data():
    try:
        # Получение всех основных категорий
        response = await make_request('https://assets.g2g.com/offer/navigation.json')

        categories_names = jmespath.search('[].cat_name.en', response)
        categories_id = jmespath.search('[].cat_id', response)
        names_and_id = dict(zip(categories_names, categories_id))
        keywords_json = await make_request('https://assets.g2g.com/offer/keyword.json')

        del names_and_id['Top up']
        # Проходимся по всем категориям и запрашиваем их под категории
        for category_name_1st, category_id in names_and_id.items():
            print('Парсинг категории -', category_name_1st)
            try:
                r = await make_request(f'https://sls.g2g.com/offer/category/{category_id}/brands')
            except aiohttp.ClientError:
                await asyncio.sleep(1)
                r = await make_request(f'https://sls.g2g.com/offer/category/{category_id}/brands')
                
            sub_category_json = r
            brands_total_offer = jmespath.search('payload.results[].total_offer', sub_category_json)
            services_id = jmespath.search('payload.results[].service_id', sub_category_json)
            brands_id = jmespath.search('payload.results[].brand_id', sub_category_json)
            # Проходимся  по 2 уровню категорий
            for brand_total_offer_2nd, service_id, brand in zip(brands_total_offer, services_id, brands_id):
                brand_name2nd = keywords_json[brand]['en']
                if isinstance(brand_total_offer_2nd, int):
                    params = {
                        'service_id': service_id,
                        'brand_id': brand,
                        'currency': 'EUR',
                        'country': 'RU',
                    }
                    try:
                        response = await make_request('https://sls.g2g.com/offer/search_result_count', params=params)
                    except aiohttp.ClientError:
                        await asyncio.sleep(1)
                        response = await make_request('https://sls.g2g.com/offer/search_result_count', params=params)
                        
                    total_result = response['payload']['total_result']
                    total_pages = total_result // 48 + (total_result % 48 > 0)

                    filter_2nd = await get_filters(service_id, brand, keywords_json)

                    for page in range(1, total_pages + 1):
                        params = {
                            'service_id': service_id,
                            'brand_id': brand,
                            'sort': 'recommended',
                            'page_size': '48',
                            'page': str(page),
                            'currency': 'EUR',
                            'country': 'RU'
                        }
                        try:
                            response = await make_request('https://sls.g2g.com/offer/search', params=params)
                        except aiohttp.ClientError:
                            await asyncio.sleep(1)
                            response = await make_request('https://sls.g2g.com/offer/search', params=params)
                            
                        search_results = response
                        # Если предложение уникальное, то значит, что 3rd уровня нет.
                        if not jmespath.search('payload.results[0].is_unique', search_results):
                            products_total_offer_3rd = jmespath.search('payload.results[].total_offer', search_results)
                            titles_3rd = jmespath.search('payload.results[].title', search_results)

                            for title_3rd, product_total_offer_3rd in zip(titles_3rd, products_total_offer_3rd):
                                ws.append([category_name_1st, brand_name2nd, brand_total_offer_2nd, str(filter_2nd), title_3rd, product_total_offer_3rd])
                        else:
                            ws.append([category_name_1st, brand_name2nd, brand_total_offer_2nd, str(filter_2nd), None, None])
                            break
                else:
                    ws.append([category_name_1st, brand_name2nd, 0, None, None, None])

        wb.save('g2g_parser.xlsx')

    except (aiohttp.ClientError, aiohttp.InvalidURL) as e:
        print(f"Произошла ошибка: {e}")
        wb.save('g2g_parser.xlsx')
        sys.exit(1)

if __name__ == '__main__':
    try:
        asyncio.run(get_data())
    except KeyboardInterrupt:
        wb.save('g2g_parser.xlsx')
        sys.exit(0)
