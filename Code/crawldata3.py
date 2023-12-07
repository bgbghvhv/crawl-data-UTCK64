import requests
from bs4 import BeautifulSoup
import openpyxl
#i=2023000002
start_row = 2
dem=0
for i in range(2023005554,2023006128):
    start_row=start_row+1
    url = 'https://xettuyen.utc.edu.vn/registration/tra-cuu-ket-qua?code='
    web=url + str(i)
    response = requests.get(web)
    soup = BeautifulSoup(response.text, 'html.parser')
    #print(soup.prettify())
#chuoi = "phần tử 1,phần tử 2,phần tử 3,phần tử 4"
    danh_sach = soup.prettify().split(",")
    workbook = openpyxl.load_workbook("data13.xlsx")
         # Chọn sheet hiện tại
    sheet = workbook.active

         # Ghi dữ liệu vào Excel
  # Hàng thứ 2 (dòng tiêu đề ở hàng 1)
    start_column = 2  # Cột thứ 2
    for i, phan_tu in enumerate(danh_sach, start=start_column):
        sheet.cell(row=start_row, column=i, value=phan_tu)
    
#         # Lưu workbook
    workbook.save("data13.xlsx")
    dem=dem+1
    print("Đã xong bản ghi", dem)
print("Hoàn thành")
