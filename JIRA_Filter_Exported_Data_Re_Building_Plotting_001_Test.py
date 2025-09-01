import tkinter as tk
from tkinter import filedialog


# ###### def open_file() Button:
# def open_file():
#     global open_file_path
#     label___Data_Re_Making_Function['text'] = ""
#     [open_file_path, open_file_name] = file_info()
#     label___open_file['text'] = open_file_name
    
#     # print(label___open_file['text'])

# ###### def file_info():
# def file_info():
#     # file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xlsx')])
#     file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xls'),('Excel File','.xlsx'),('Excel File','.csv'),('All File','.*')])
#     file_path = file_path.replace('\\', '/')

#     index = file_path.rfind("/")
#     file_name = file_path[index+1:len(file_path)]

#     return file_path, file_name


# ###### def Data_Re_Making_Bottom() Button:
# def Data_Re_Making_Bottom():
#     if label___open_file['text'] == 'Please Select File' :
#         label___Data_Re_Making_Function['text'] = "Error : Please Select File"
#         # print(label___open_file['text'])
#     else :
#         label___Data_Re_Making_Function['text'] = "Re-Making Now"
#         open_file_path
#         # print(label___open_file['text'])
#         ##### ploting
#         Data_Re_Making_Function(open_file_path)
#         label___Data_Re_Making_Function['text'] = "Finish"

###### def find_cell_info():
def find_cell_info(input_book, target_text):
    found_address = 0
    index_front = 0
    index_rear = 0

    found_cell = input_book.sheets(1).used_range.api.Find(What=target_text)
    found_address = found_cell._inner.Address

    index_front = found_address.find("$")
    index_rear = found_address.rfind("$")

    found_address = found_address[index_front+1:index_rear]

    return found_address

###### def Data_Re_Making_Function(input_path):
# def Data_Re_Making_Function(input_path):
import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

##### 파일 불러오기
input_path_1 = r'C:\Users\namho.chung\Desktop\업무\STUDY\JIRA Design Change Export from Filter\Test\20250826 1527 SearchRequest-20169_Re-Making.xlsx' # r'brabra' Raw String
input_path_2 = r'C:\Users\namho.chung\Desktop\업무\STUDY\JIRA Design Change Export from Filter\Test\20250828 0749 SearchRequest-20169_Re-Making.xlsx' # r'brabra' Raw String 
input_path_3 = r'C:\Users\namho.chung\Desktop\업무\STUDY\JIRA Design Change Export from Filter\Test\20250901 1051 SearchRequest-20169_Re-Making.xlsx' # r'brabra' Raw String 


##### 불러온 파일 정리1
# xw.App(visible = False)
xw.App(visible = True)
book_1 = xw.Book(input_path_1)
book_2 = xw.Book(input_path_2)
book_3 = xw.Book(input_path_3)
# row_size = book.sheets(1).range('A1').current_region.last_cell.row
# column_size = book.sheets(1).range('A1').current_region.last_cell.column


# np_excel_data = book.sheets(1).range("A1").current_region.options(np.array).value
# del np_excel_data
pd_excel_data_1 = book_1.sheets["pivot 부서별-상태-개수"].range("A1").current_region.options(pd.DataFrame).value
pd_excel_data_2 = book_2.sheets["pivot 부서별-상태-개수"].range("A1").current_region.options(pd.DataFrame).value
pd_excel_data_3 = book_3.sheets["pivot 부서별-상태-개수"].range("A1").current_region.options(pd.DataFrame).value

np_excel_data_1 = book_1.sheets["pivot 부서별-상태-개수"].range("A1").current_region.options(np.array).value
np_excel_data_2 = book_2.sheets["pivot 부서별-상태-개수"].range("A1").current_region.options(np.array).value
np_excel_data_3 = book_3.sheets["pivot 부서별-상태-개수"].range("A1").current_region.options(np.array).value

book_1.app.kill()

# print(pd_excel_data_1)
# print(pd_excel_data_2)
# print(pd_excel_data_3)

# print(np_excel_data_1)
# print(np_excel_data_2)
# print(np_excel_data_3)

xpos_1 = list(range(1, len(pd_excel_data_1.index)+1, 1))
ypos_1 = list(range(0,3,1))


ratio  = len(ypos_1) / len(xpos_1)
thickness = 0.8

print(ratio)
print(thickness)

xpos_1, ypos_1 = np.meshgrid(xpos_1, ypos_1)
xpos_1 = xpos_1.ravel()
ypos_1 = ypos_1.ravel()

zpos = 0

dx = 0.5
dy = 0.5
dz_1 = pd_excel_data_1['1그룹'].to_list()
dz_2 = pd_excel_data_2['1그룹'].to_list()
dz_3 = pd_excel_data_3['1그룹'].to_list()

dz = dz_1 + dz_2 + dz_3


##### 3d figure
fig = plt.figure()
ax = fig.add_subplot(projection='3d')
# ax.bar3d(xpos_1, ypos_1, zpos, dx, dy, dz, zsort='average')
# ax.bar3d(xpos_1, ypos_1, zpos, dx, dy, dz)
print(xpos_1)
print(ypos_1)
print(dz)



ax.bar3d(xpos_1, ypos_1, 0, thickness, ratio*thickness, dz)
# ax.bar3d(xpos_1, ypos_1, 0, 0.8, 0.8, dz)
ax.set_xlabel('X Status')
ax.set_ylabel('Y Date')
ax.set_zlabel('Z EA')

plt.show()

print('TEST')


# ##### 창 만들기 #####
# root = tk.Tk()
# # width, height = 500, 25 # 창 크기 값 설정
# # get_path = None
# root.geometry("400x100") # 창 크기 설정
# root.resizable(True, True) # 크기 조정 가능 여부
# root.title('JIRA Filter Exported Data Re-Building') # 창 제목 설정

# # open_file 
# button___open_file = tk.Button(root, text="1. Select File", command=open_file)
# button___open_file.pack()

# label___open_file = tk.Label(root, text="Please Select File")
# label___open_file.pack()

# # Data_Re_Making_Function
# button___Data_Re_Making_Function = tk.Button(root, text="2. Data Re-Making", command=Data_Re_Making_Bottom)
# button___Data_Re_Making_Function.pack()

# label___Data_Re_Making_Function = tk.Label(root, text="")
# label___Data_Re_Making_Function.pack()

# root.mainloop()
