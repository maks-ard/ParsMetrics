"""Редактирует файл с id, если что то меняется в таблице"""
import json
import pandas as pd


def delete_element_in_table(key_after: int):
    with open(r'../data/ids.json', 'r') as file:
        ids = json.load(file)
        result = {}
    for key, item in ids.items():
        key = int(key)
        if key > key_after:
            key += 1
            result[str(key)] = item
        else:
            result[str(key)] = item

    with open(r'../data/ids.json', 'w') as x:
        json.dump(result, x, indent=4, ensure_ascii=False)


def test():
    daterange = pd.date_range('2022-02-24', '2022-03-13')
    for single_date in daterange:
        print(str(single_date.strftime("%Y-%m-%d")))


if __name__ == '__main__':
    # delete_element_in_table(66)
    test()
