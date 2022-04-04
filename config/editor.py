"""Редактирует файл с id, если что то меняется в таблице"""
import json


class Editor:
    def __init__(self):
        self.path = r'data/ids.json'

    def delete_element_in_table(self, id_goal):
        with open(self.path, 'r') as file:
            ids = json.load(file)
        try:
            del ids[id_goal]
        except KeyError:
            pass
        res = {}
        for key, goal in ids.items():
            key = int(key)
            if key > int(id_goal):
                key -= 1
            res[key] = goal
        with open(self.path, 'w') as result:
            json.dump(res, result, indent=4, ensure_ascii=False)


if __name__ == '__main__':
    id_goal = input("Какую цель необходимо удалить?: ")
    edit = Editor()
    edit.delete_element_in_table(id_goal=id_goal)
