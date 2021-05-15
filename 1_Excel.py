from openpyxl import Workbook

wb = Workbook() #새 워크북 생성

ws = wb.active # 현재 활성화된 sheet 가져온다. 

ws.title = "SJSheet" #sheet의 이름을 변경 
wb.save("sample.xlsx")
wb.close()