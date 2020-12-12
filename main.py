
import xlsxwriter as xs

workbook = xs.Workbook("Student_Information.xlsx")
worksheet = workbook.add_worksheet("Student Details")

Font_Format = workbook.add_format()
font_format_row = workbook.add_format()
Font_Format.set_font_size(15)
Font_Format.set_bold()
Font_Format.set_bg_color('yellow')
worksheet.set_column('A:C',30)
font_format_row.set_size(13)


worksheet.write('A1','STUDENT NAMES', Font_Format)
worksheet.write('B1','ADDRESS', Font_Format)
worksheet.write('C1','PHONE No.', Font_Format)

# name = input("Enter your Name: ")
# worksheet.write('A2', name)

n = int(input("How many Fields you want to put: "))
content = []
content_addr = []
content_phone = []

for i in range(0,n):
    print("Enter Names with Surname: ")
    names = input().title()
    print("Enter Address or City: ")
    address = input().title()
    print("Enter Phone Number: ")
    phone = int(input())
    content.append(names)
    content_addr.append(address)
    content_phone.append(phone)

# print(content_addr)

# code for adding names cells
row = 1
col = 0
for namesCol in content:
    worksheet.write(row,col,namesCol,font_format_row)
    row += 1

# code for adding address cells
row = 1
col = 1
for addrCol in content_addr:
    worksheet.write(row,col,addrCol,font_format_row)
    row += 1

# code for adding phone cells
row = 1
col = 2
for phnCol in content_phone:
    worksheet.write(row,col,phnCol,font_format_row)
    row += 1

print("Great! Now Check your project directory!")
workbook.close()
