from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.insert_rows(8) # 8 번째 줄에 빈 칸 삽입
#ws.insert_rows(8,5) # 8 번째 줄에 빈칸 5개 삽입

ws.insert_cols(2,3) #B열에 빈 열 3개 추가하기 

wb.save("sample_insert_cols.xlsx")