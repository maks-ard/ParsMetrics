from urllib import response
import openpyxl, requests, datetime


filename = 'metrics.xlsx'
url = 'https://metrika.yandex.ru/stat/conversion_rate?no_robots=1&robots_metric=1&cross_device_attribution=1&cross_device_users_metric=0&group=dekaminute&period=yesterday&accuracy=1&id=19405381'



def parse_metrics():
	headers = {
		"Host": "metrika.yandex.ru",
		"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
		"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:96.0) Gecko/20100101 Firefox/96.0"
		# 'Cookie': 'yandexuid=2767614281641979789; is_gdpr=0; is_gdpr_b=CMTYOBC4YSgC; _yasc=smTeHW69IDAhDs3iFjiFbUp6ad2Jx15ZAyDKtvJqKdvwquTG7z6FYNYDrQBQi0yppLs=;\
		# 			i=L68r2Uld7G1sR/pTqdIphPcKDRg2fQZuESQc9tC26zsXNfkViZhfEIQbZBsUUPNF3kQXr8YUaOKwIiXALcggh2dffYQ=; \
		# 			yabs-frequency=/5/00020000001SlWDY/ujohYroPR7uOHU2gPGBOHWmaSnX5W0TU-7bLgZvS_qJy____XcGPErUFz6SOHK05XSAJPRzFLnX5OBRYV0tZjMrp64KW/; \
		# 			yp=1676448724.p_sw.1644912724#1676359957.p_cl.1644823957#1645016376.mcv.0#1645016376.mct.null#1645016376.mcl.1hqjiwu#1645016376.szm.1%3A1920x108…=0; \
		# 			Session_id=3:1644914890.5.0.1644826261254:MurAPg:12.1.2:1|1520876748.88629.2.2:88629|3:248119.482589.J145mBO3g3y7y97mJmtNPOnlmSY; \
		# 			sessionid2=3:1644914890.5.0.1644826261254:MurAPg:12.1.2:1|1520876748.88629.2.2:88629|3:248119.482589.J145mBO3g3y7y97mJmtNPOnlmSY; \
		# 			L=CQJpVgQGSlJCVmV5bUppVnx1WwxSUmd/Kz1cNxUiGC4yNg==.1644914890.14889.387869.d55951f1f0db780ba4efc0b218bf388c; yandex_login=ardeev.max; \
		# 			mda=0; yandex_gid=20; my=YwA=; computer=1; XcfPaDInQpzKj=1; 71f6f4a784872450ca476fdbfe296e9d=1; _ymfc=p:yesterday'
		# 'login': 'ardeev.max@yandex.ru',
		# 'password': 'Thebesthimik2021'
	}
	response = requests.get(url, headers=headers, cookies=)
	print(response.status_code)

	with open('index.html', 'w', encoding='utf-8') as file:
		file.write(response.text)

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
	parse_metrics()