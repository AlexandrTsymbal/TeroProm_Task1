import pandas as pd
from fuzzywuzzy import process


def get_general_info(supl_path: str, tree_path: str):
    """
    Основная функция, получаящая данные о товаре (наименование\категория постащика)
    и составляющая связь товар - категория

    :param supl_path: Путь до "Данных поставщика"
    :param tree_path: Путь до "Дерева категорий"
    :return:
    """
    dict_pres_type = {}
    name_to_pres = {}
    gen_info = pd.read_excel(supl_path, usecols=[1, 3])
    lst, categories = get_list_types(tree_path)

    lst_set = set(lst)

    for row in gen_info.itertuples(index=True):
        type_product = row[2].rsplit('/', 1)
        product_name = row[1]

        name_to_pres[product_name] = type_product[-1]

        if type_product[-1] in dict_pres_type:
            res_name, score_name = search_in_type_tree(' '.join(product_name.split()[:3]), lst_set)
            if dict_pres_type[type_product[-1]][1] < score_name:
                dict_pres_type[type_product[-1]] = (res_name, score_name)
        else:
            res_cat, score_cat = search_in_type_tree(type_product[-1], lst_set)
            res_name, score_name = search_in_type_tree(' '.join(product_name.split()[:3]), lst_set)
            dict_pres_type[type_product[-1]] = max(
                (res_cat, score_cat), (res_name, score_name), key=lambda x: x[1]
            )

        print(row[0])
    create_new_table(dict_pres_type, categories, name_to_pres)


def get_list_types(tree_path) -> tuple:
    """
    Получаем список всех возможных типов и
    Связи вида тип-категория

    :param tree_path: Путь до "дерева категорий"
    :return:
    """
    tree_data = pd.read_excel(tree_path, dtype=str)
    categories = tree_data.iloc[:, [0, 1, 2]].values.tolist()
    return tree_data.iloc[:, 2].tolist(), categories


def search_in_type_tree(presumptive: str, lst_colums: set):
    """
    Поиск по расстоянию Левинштейна

    :param presumptive: Предполагаемая категория\наименование товара
    :param lst_colums:Список возможных категорий
    :return: Кортеж (Категория, вероятность совпадения)
    """
    best_match = process.extractOne(presumptive, lst_colums)  # RapidFuzz
    return best_match


def create_new_table(dict_cat: dict, list_types: list, names: dict):
    """
    Создание новой таблице с найдеными категориями

    :param dict_cat: Словарь предполгаемой категории : найденной категории
    :param list_types: Список связей категория - тип
    :param names:  Словарь наименование : категория
    :return:
    """
    rows = []
    for prod, type_key in names.items():
        type_prod = dict_cat[type_key][0]
        for types in list_types:
            if type_prod in types:
                rows.append({
                    'Главная категория': types[2],
                    'Дочерняя категория': types[1],
                    'Тип товара': type_prod,
                    'Товар': prod
                })

    new_table = pd.DataFrame(rows)
    new_table.to_excel(r'files\output.xlsx', index=False)


if __name__ == '__main__':
    supplier_data = r'..\files\Данные поставщика.xlsx'
    tree_categories = r'..\files\Дерево категорий.xlsx'

    get_general_info(supplier_data, tree_categories)
