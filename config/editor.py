"""Редактирует файл с id, если что то меняется в таблице"""
import json


class Editor:
    def __init__(self):
        self.path = r'data/ids.json'

    def delete_element_in_table(self, id_goal):
        with open(self.path, 'r') as file:
            ids = json.load(file)
        del ids[id_goal]
        res = {}
        for key, goal in ids.items():
            key = int(key)
            if key > int(id_goal):
                key -= 1
            res[key] = goal
        with open(self.path, 'w') as result:
            json.dump(res, result, indent=4, ensure_ascii=False)


edit = Editor()
edit.delete_element_in_table(id_goal='10')
