#셀레늄 웹드라이버
from selenium import webdriver

#웹드라이버 객체 생성시 수반될 서비스나 옵션
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

#선택자 및 키보드 입력
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


# 웹 드라이버 매니저 모듈
from webdriver_manager.chrome import ChromeDriverManager

# 기타모듈
import time

#by 선언
from selenium.webdriver.common.by import By

#서비스 변수 생성 
customService = Service(ChromeDriverManager().install()) 
#옵션 변수 생성 
customOption = Options() 

#드라이버 객체 생성 
browser = webdriver.Chrome(service = customService, options = customOption)



URL = "https://www.naver.com/"

browser.get(URL)
browser.implicitly_wait(10)

#메일 값 획득
temp = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/div[5]/ul/li[1]/a/span[2]').text
print(temp)


# #만약 조금 상위 레벨에서 진행
# temp = browser.find_element(By.XPATH, '//*[@id="shortcutArea"]/ul/li[1]/a').text
# print(temp)


#send_keys
browser.find_element(By.XPATH, '//*[@id="query"]').send_keys('2025-06-15')
time.sleep(3)


#버튼 클릭
browser.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[2]/div[2]/div/div/div[1]/div/a').click()

time.sleep(3)


#엑셀파일 생성
xlsxFile = openpyxl.Workbook()

#생성한 파일에서 시트 생성
xlsxSheet = xlsxFile.active


#시트 특정 셀에 데이터 입력
xlsxSheet.cell(row = 1, column = 1).value = "hi"
#find_element().text 로 찾은 값을 넣으면 됨


#저장
xlsxFile.save('result.xlsx')

