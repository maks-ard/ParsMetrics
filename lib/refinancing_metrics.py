import pandas as pd
import openpyxl
import traceback
from datetime import datetime, timedelta
from openpyxl.styles.numbers import BUILTIN_FORMATS

from common import BaseExcel, api_yandex_async


class RefinancingExcel(BaseExcel):
    def __init__(self, filename=None):
        super(RefinancingExcel, self).__init__()
        self.filename = self.get_filepath("CR Перекредитование в ЛК") if filename is None else filename
        self.book = openpyxl.load_workbook(self.filename)

    @staticmethod
    def get_row_for_write(sheet):
        """Получить ряд для записм данных"""
        rows = sheet.iter_rows(max_col=1)
        result = ()
        for row in rows:
            for cell in row:
                result = (cell.value, cell.row)
                if cell.value is None:
                    return cell.offset(row=-1).value, cell.offset(row=-1).row
        return result

    def get_ids_refinancing(self, is_comment=True):
        """Получить колонки для записи"""
        result = {}
        for name in self.book.sheetnames:
            sheet = self.book[name]
            result[name] = {} if is_comment else []

            for row in sheet.iter_rows(max_row=1, min_col=2, max_col=sheet.max_column):
                for cell in row:
                    if is_comment:
                        if cell.comment is not None:
                            result[name].update({cell.comment.text.split("\n")[1]: cell.column_letter})
                    else:
                        if cell.comment is None:
                            result[name].append(cell.column)

        return result

    def do_offset_formulas(self, sheet, data, row):
        for col in data:
            cell = sheet.cell(row=row, column=col)
            cell_offset = cell.offset(row=-1)
            cell_offset_row = cell_offset.row
            try:
                if type(cell_offset.value) is str:
                    if cell_offset.value[0] == "=":
                        result = cell_offset.value.replace(str(cell_offset_row), str(row))

                        sheet[cell.coordinate] = result

                        if "РСД" not in cell_offset.value:
                            sheet[cell.coordinate].number_format = BUILTIN_FORMATS[9]

            except Exception:
                self.logger.error(traceback.format_exc())

    def main(self, filename=None):
        data = self.get_ids_refinancing()
        data_formulas = self.get_ids_refinancing(is_comment=False)

        for name, ids in data.items():
            sheet = self.book[name]
            last_row = self.get_row_for_write(sheet)

            first_date = str(last_row[0])
            row = int(last_row[1])
            yesterday = (datetime.now() - timedelta(days=1)).strftime('20%y-%m-%d')
            daterange = pd.date_range(first_date.split(" ")[0], yesterday)

            self.logger.info(f"{last_row}:{yesterday}")

            if first_date.split(" ")[0] != datetime.now().strftime('20%y-%m-%d'):
                for date in daterange:
                    date_format = str(date.strftime("%Y-%m-%d"))
                    metrics = api_yandex_async.main(ids, date1=date_format, date2=date_format, need_users=False)

                    sheet[f"A{row + 1}"] = date + timedelta(days=1)

                    for col, goal in metrics.items():
                        if goal != "" and col != "date":
                            sheet[str(col) + str(row)] = int(goal)

                    self.do_offset_formulas(sheet, data_formulas[name], row)
                    row += 1

        self.book.save(self.filename)
