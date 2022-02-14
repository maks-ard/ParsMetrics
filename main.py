import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet import worksheet
import datetime

filename = 'metrics.xlsx'


def edit_file():
	book = openpyxl.load_workbook(filename=filename)
	sheet = book.active

	all_col = {
		'01' : 'B', '02' : 'D','03' : 'F','04' : 'H','05' : 'J','06' : 'L','07' : 'N','08' : 'P','09' : 'R','10' : 'T',
		'11' : 'V','12' : 'X','13' : 'Z','14' : 'AB','15' : 'AD','16' : 'AF','17' : 'AH','18' : 'AJ','19' : 'AL','20' : 'AN',
		'21' : 'AP','22' : 'AR','23' : 'AT','24' : 'AV','25' : 'AX','26' : 'AZ','27' : 'BB','28' : 'BD','29' : 'BF','30' : 'BH','31' : 'BJ'
	}
	x = [3, 4, 6, 7, 8, 11, 12, 13, 14, 16, 17, 18, 19, 21, 22, 23, 24, 26, 27, 28, 29, 31, 32, 33, 34, 
		35, 37, 39, 40, 41, 43, 44, 45, 46, 47, 49, 50, 53, 53, 54, 57, 58, 60, 61, 62, 63, 64, 65, 66]

	now_day = datetime.datetime.today().strftime("%d")
	col = all_col[now_day]

	for item in range(len(x)):
		sheet[col + str(x[item])] = 100

	book.save(filename=filename)

if __name__=='__main__':
	edit_file()