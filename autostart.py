from config import writing_ecxel


def get_params():
    writing_ecxel.edit_file()


if __name__ == '__main__':
    get_params()

    import os
    os.startfile(r"C:\Users\m.ardeev\OneDrive - ООО Микрокредитная компания «Центрофинанс Групп»\Документы\Метрики2022КОПИЯ.xlsx")
