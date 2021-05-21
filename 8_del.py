from openpyxl import load_workbook


wb = load_workbook("sample.xlsx")
ws = wb.active

#ws.delete_cols(2) #8번째 줄에 있는 데이터 삭제 

#ws.delete_rows(3) # 8 번째 줄에 있는 7번 ㄷ학생 데이터 삭제

ws.delete_cols(2,2)
wb.save("sample_del_cols1.xlsx")