import pandas as pd
import traceback
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from progress.bar import IncrementalBar

from common import BaseExcel, api_yandex_async


class RefinancingExcel(BaseExcel):
    def __init__(self, filename=None):
        super().__init__()
        self.filename = self.get_filepath("CR Перекредитование в ЛК") if filename is None else self.get_filepath(
            filename)
        self.book = openpyxl.load_workbook(self.filename)

    @staticmethod
    def get_row_for_write(sheet: Worksheet) -> tuple:
        """Получить последную дату и ряд"""
        return [(row[0].value, row[0].row) for row in sheet.iter_rows(max_col=1) if row[0].value is not None][-1]

    def get_ids_refinancing(self, is_comment=True) -> dict[dict] | dict[list]:
        """
        is_comment: True - Получить id целей и их колонки dict[dict]
        is_comment: False - Получить колонки для формул dict[list]
        """
        result = {name: {} if is_comment else [] for name in self.book.sheetnames}

        for name in self.book.sheetnames:
            sheet: Worksheet = self.book[name]

            for row in sheet.iter_rows(max_row=1, min_col=2, max_col=sheet.max_column):
                for cell in row:

                    if is_comment and cell.comment is not None:
                        comment = cell.comment.text.split("\n")[1]
                        result[name].update({comment: cell.column})

                    if not is_comment and cell.comment is None and cell.value is not None:
                        result[name].append(cell.column)
        return result

    def do_offset_formulas(self, sheet: Worksheet, data: list, row: int):
        for col in data:

            try:
                cell: Cell = sheet.cell(row=row, column=col)
                cell_offset = cell.offset(row=-1)

                if isinstance(cell_offset.value, str) and cell_offset.value.startswith("="):
                    sheet[cell.coordinate] = cell_offset.value.replace(str(cell_offset.row), str(row))

                    if "РСД" not in cell_offset.value:
                        sheet[cell.coordinate].number_format = BUILTIN_FORMATS[9]

            except Exception:
                print("Ошибка в записи формул!")
                self.logger.error(traceback.format_exc())

    def main(self):
        bar = IncrementalBar("Выгрузка перекредитования:", max=len(self.book.sheetnames))

        goals = self.get_ids_refinancing()
        formulas = self.get_ids_refinancing(is_comment=False)

        for name, ids in goals.items():
            sheet: Worksheet = self.book[name]
            last_row = self.get_row_for_write(sheet)
            first_date: datetime = last_row[0]

            if first_date.date() != datetime.now().date():
                daterange = pd.date_range(first_date, self.get_yesterday.strftime('%Y-%m-%d'))
                row = last_row[1]

                for date in daterange:
                    sheet.cell(row=row + 1, column=1, value=date + timedelta(days=1))
                    sheet[f"A{row + 1}"].number_format = "DD.MM.YYYY"

                    self.do_offset_formulas(sheet, formulas[name], row)

                    date_format = date.strftime("%Y-%m-%d")
                    metrics = api_yandex_async.main(ids, date1=date_format, date2=date_format, need_users=False)

                    self.logger.info(f"{name}: {metrics}")

                    for col, goal in metrics.items():
                        if goal != "" and col != "date":
                            sheet.cell(row=row, column=col, value=int(goal))

                    row += 1

            bar.next()

        self.book.save(self.filename)

        bar.finish()

