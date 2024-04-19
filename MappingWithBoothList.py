import gspread
from gspread.cell import rowcol_to_a1
from gspread.utils import ValueRenderOption
import time

def checkSpecialBooth(BoothNumber: str):
	specialbooth_code_list = ['Vir', 'Cre', 'Psm', 'Adt', 'AZ', 'Voc']
	for code in specialbooth_code_list:
		if code in BoothNumber:
			return True
		else:
			continue
	return False

def SetLinkToMap(BoothListSheet: gspread.Worksheet, BoothMapSheet: gspread.Worksheet, BoothNumber: str):
	"""
	부스 번호 셀과 부스 지도에서의 해당 부스 위치 셀을 서로 링크합니다.

	- 매개 변수
	:param BoothNumber: 서로 링크할 부스 번호
	"""

	BoothNumberCell_Data = BoothListSheet.find(BoothNumber)

	BoothNumber_splited = BoothNumber.replace("\n", " ").split(", ") if ',' in BoothNumber else [BoothNumber]

	# key => 지도에서의 해당 부스의 a1 위치 값, value => 부스 위치에서의 a1 위치 값
	BoothLocations = []
	for Number in BoothNumber_splited:
		temp = Number.replace(' ', '\n') if checkSpecialBooth(Number) == True else Number
		MapLocationData = BoothMapSheet.find(temp)
		BoothLocations.append(rowcol_to_a1(MapLocationData.row, MapLocationData.col))

		map_value = f'TEXTJOIN(CHAR(10), 0, "{Number.split(" ")[0]}", "{Number.split(" ")[1]}")' if checkSpecialBooth(Number) == True else f'"{Number}"'
			
		BoothMapSheet.update_acell(rowcol_to_a1(MapLocationData.row, MapLocationData.col),
						  		f'=HYPERLINK("#gid={BoothListSheet.id}&range={rowcol_to_a1(BoothNumberCell_Data.row, BoothNumberCell_Data.col)}", {map_value})')

	BoothListSheet.update_acell(rowcol_to_a1(BoothNumberCell_Data.row, BoothNumberCell_Data.col),
						  		f'=HYPERLINK("#gid={BoothMapSheet.id}&range={BoothLocations[0]}:{BoothLocations[len(BoothLocations) - 1]}", "{BoothNumber}")')

def printDebug(tag: str, vari: any):
	print(f"{tag} : {vari}")

gc : gspread.client.Client = None
sheet : gspread.spreadsheet.Spreadsheet = None

sheetId = "1TzNBg9FTmXkrpBtpxdBiI5hNdJn76NPgtMkmGAwYinM"
sheetNumber = 0
MapSheetNumber = 6

special_booth_zone_name_inkorean = ['버츄올스타', '크리에스타', '동방특별존', '어른의 특별존', '보카스타', '종합', '초대형 서클', '기타']

Is_specialBoothTitle = False

gc = gspread.service_account()
sheet_ = gc.open_by_key(sheetId)
boothlist_Sheet = sheet_.get_worksheet(sheetNumber)
BoothMapSheet = sheet_.get_worksheet(MapSheetNumber)

boothNumber_list = boothlist_Sheet.get(f'B:B', value_render_option=ValueRenderOption.formatted)
boothNumber_list_completed: list[str] = []
for i in range(2, len(boothNumber_list)):
	if len(boothNumber_list[i]) == 0:
		continue
	for j in range(0, len(special_booth_zone_name_inkorean)):
		if special_booth_zone_name_inkorean[j] in boothNumber_list[i][0]:
			#print(f"Is Is_specialBoothTitle is set to True")
			Is_specialBoothTitle = True
			break
	if Is_specialBoothTitle == True:
		Is_specialBoothTitle = False
		continue
	else:
		boothNumber_list_completed.append(boothNumber_list[i][0])
	

printDebug("boothNumber_List", boothNumber_list)
printDebug("boothNumber_list_completed", boothNumber_list_completed)

for boothnumber in boothNumber_list_completed:
	SetLinkToMap(boothlist_Sheet, BoothMapSheet, boothnumber)
	time.sleep(2)