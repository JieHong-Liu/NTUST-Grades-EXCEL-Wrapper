from bs4 import BeautifulSoup
import openpyxl

path = 'score.html'
htmlfile = open(path, 'r', encoding='utf-8')
htmlhandle = htmlfile.read()
soup = BeautifulSoup(htmlhandle, 'lxml')
tables = soup.find_all('table')

tr_list = tables[1].find_all('tr')
term_list = []  # 學期
course_names_list = []  # 課程名稱
credits_list = []  # 學分數
grades_list = []  # 成績

for tr in tr_list:
    td_list = tr.find_all('td')
    for td in range(len(td_list)):
        if(td == 1):
            term_list.append(td_list[1].text.strip())
        elif(td == 3):
            course_names_list.append(td_list[3].text.strip())
        elif(td == 4):
            credits_list.append(td_list[4].text.strip())
        elif(td == 5):
            grades_list.append(td_list[5].text.strip())

# 利用 Workbook 建立一個新的工作簿
workbook = openpyxl.Workbook()

# 取得第一個工作表
sheet = workbook.worksheets[0]

# initialize
write_list = ['學期年度', '課程名稱', '學分數', '成績']
c_list = ['A1', 'B1', 'C1', 'D1']
for i in range(len(write_list)):
    sheet[c_list[i]] = write_list[i]

alphabet = ['A', 'B', 'C', 'D']

for i in range(2, 2+len(course_names_list)):
    for j in range(0, 4):
        if (j == 0):
            sheet[alphabet[j]+str(i)] = term_list[i-2]
        elif(j == 1):
            sheet[alphabet[j]+str(i)] = course_names_list[i-2]
        elif(j == 2):
            sheet[alphabet[j]+str(i)] = credits_list[i-2]
        elif(j == 3):
            sheet[alphabet[j]+str(i)] = grades_list[i-2]

# 儲存檔案
workbook.save('grades.xlsx')
