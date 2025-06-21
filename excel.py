#액셀 모듈 임포트
import openpyxl



#엑셀파일 생성
xlsxFile = openpyxl.Workbook()

#생성한 파일에서 시트 생성
xlsxSheet = xlsxFile.active


#시트 특정 셀에 데이터 입력
for i in range(10):
    xlsxSheet.cell(row = i + 1, column = 1).value = "hi"
#find_element().text 로 찾은 값을 넣으면 됨


#저장
xlsxFile.save('result.xlsx')

