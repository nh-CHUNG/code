import tkinter as tk
from tkinter import filedialog
from functools import partial


###### Preload Funcitons ######
###### def file_info():
def file_info():
    # file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xlsx')])
    file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xls'),('Excel File','.xlsx'),('Excel File','.csv'),('All File','.*')])
    file_path = file_path.replace('\\', '/')

    index = file_path.rfind("/")
    file_name = file_path[index+1:len(file_path)]

    return file_path, file_name


###### def find_cell_info():
# def find_cell_info(input_book, target_text):
#     import xlwings as xw
#     found_address = 0
#     index_front = 0
#     index_rear = 0

#     found_cell = input_book.sheets(1).used_range.api.Find(What=target_text)
#     found_address = found_cell._inner.Address

#     index_front = found_address.find("$")
#     index_rear = found_address.rfind("$")

#     found_address = found_address[index_front+1:index_rear]

#     return found_address


###### Defination of Button ######
###### def Data_Accumulate_Button():
# def Data_Accumulate_Button():
#     if label___open_file['text'] == 'Please Select File' :
#         label___Data_Accumulate['text'] = "Error : Please Select File"
#         # print(label___open_file['text'])
#     else :
#         label___Data_Accumulate['text'] = "Data Accumulating Now"
#         # print(label___open_file['text'])
#         ##### ploting
#         Data_Accumulate_Function(OPEN_FILE_PATH)
#         label___Data_Accumulate['text'] = "Finish"



###### def Data_Accumulate_Function(input_path):
# def Data_Accumulate_Function(input_path):
#     import xlwings as xw
#     import pandas as pd
#     import matplotlib.pyplot as plt
#     global PD_EXCEL_DATA

#     ##### 파일 불러오기
#     # input_path_1 = r'C:\Users\namho.chung\Desktop\업무\STUDY\JIRA Design Change Export from Filter\Test\20250826 1527 SearchRequest-20169_Re-Making.xlsx' # r'brabra' Raw String

#     ##### 불러온 파일 정리1
#     # xw.App(visible = False)
#     xw.App(visible = True)
#     book = xw.Book(input_path)
#     # row_size = book.sheets(1).range('A1').current_region.last_cell.row
#     # column_size = book.sheets(1).range('A1').current_region.last_cell.column

#     PD_EXCEL_DATA = book.sheets["pivot 부서별-상태-개수"].range("A1").current_region.options(pd.DataFrame).value
#     print(PD_EXCEL_DATA)
#     book.app.kill()
#     del book



###### def Data_Ploting_Function():
# def Data_Ploting_Function():
#     import numpy as np
#     import matplotlib.pyplot as plt

#     xpos = list(range(1, len(PD_EXCEL_DATA.index)+1, 1))
#     ypos = list(range(0,3,1))
#     xpos, ypos = np.meshgrid(xpos, ypos)

#     xpos = xpos.ravel()
#     ypos = ypos.ravel()
#     zpos = 0

#     ratio  = len(ypos) / len(xpos)
#     thickness = 0.8

#     dx = thickness
#     dy = ratio * thickness
#     dz = PD_EXCEL_DATA['1그룹'].to_list()
#     # dz_2 = pd_excel_data_2['1그룹'].to_list()
#     # dz_3 = pd_excel_data_3['1그룹'].to_list()
#     # dz = dz_1 + dz_2 + dz_3


#     ##### 3d figure
#     fig = plt.figure()
#     ax = fig.add_subplot(projection='3d')
#     # ax.bar3d(xpos, ypos, zpos, dx, dy, dz, zsort='average')
#     ax.bar3d(xpos, ypos, zpos, dx, dy, dz)

#     ax.set_xlabel('X Status')
#     ax.set_ylabel('Y Date')
#     ax.set_zlabel('Z EA')

#     plt.show()

#     print('TEST')


##### 창 만들기 #####
root = tk.Tk()
root.geometry("450x200") # 창 크기 설정 가로x세로
root.resizable(True, True) # 크기 조정 가능 여부
root.title('JIRA Filter Re-Building Data Ploting') # 창 제목 설정

# 확장 가능한 행과 열 설정
# root.columnconfigure(0, weight=1) # 0번 열↕ 확장 가능
# root.rowconfigure(0, weight=1)    # 0번 행↔ 확장 가능
root.columnconfigure(1, weight=1) # 1번 열↕ 확장 가능
# root.rowconfigure(1, weight=1)    # 1번 행↔ 확장 가능


# 1. Select File
Label___open_file = tk.Label(root, text="1. Select File")
Label___open_file.grid(row=0, column=0, columnspan=3, sticky='EW')


# 1-1. Select File Button
###### def open_file_button():
def open_file_button(Text___open_file):
    open_file_path = ''
    open_file_name = 'Please Select File...'
    [open_file_path, open_file_name] = file_info()
    if open_file_path == '' or open_file_name == '':
        Text___open_file['text'] = 'Please Select File...'   




Label___open_file_A = tk.Label(root, text="File A : ")
Label___open_file_A.grid(row=1, column=0, sticky='E')
Text___open_file_A = tk.Label(root, text="Please Select File...", height = 1, width = 45, bg="lightblue")
Text___open_file_A.grid(row=1, column=1, padx=5, pady=5, sticky='WE')
button___open_file_A  = tk.Button(root, text="Select...", command=lambda: open_file_button(Text___open_file_A))
button___open_file_A.grid(row=1, column=2, padx=5, pady=5, sticky='W')


Label___open_file_B = tk.Label(root, text="File B : ")
Label___open_file_B.grid(row=2, column=0, sticky='E')
Text___open_file_B = tk.Label(root, text="Please Select File...", height = 1, width = 45, bg="lightblue")
Text___open_file_B.grid(row=2, column=1, padx=5, pady=5, sticky='WE')
button___open_file_B  = tk.Button(root, text="Select...", command=lambda: open_file_button(Text___open_file_B))
button___open_file_B.grid(row=2, column=2, padx=5, pady=5, sticky='W')


Label___open_file_C = tk.Label(root, text="File C : ")
Label___open_file_C.grid(row=3, column=0, sticky='E')
Text___open_file_C = tk.Label(root, text="Please Select File...", height = 1, width = 45, bg="lightblue")
Text___open_file_C.grid(row=3, column=1, padx=5, pady=5, sticky='WE')
button___open_file_C  = tk.Button(root, text="Select...", command=lambda: open_file_button(Text___open_file_C))
button___open_file_C.grid(row=3, column=2, padx=5, pady=5, sticky='W')


Label___open_file_D = tk.Label(root, text="File D : ")
Label___open_file_D.grid(row=4, column=0, sticky='E')
Text___open_file_D = tk.Label(root, text="Please Select File...", height = 1, width = 45, bg="lightblue")
Text___open_file_D.grid(row=4, column=1, padx=5, pady=5, sticky='WE')
button___open_file_D  = tk.Button(root, text="Select...", command=lambda: open_file_button(Text___open_file_D))
button___open_file_D.grid(row=4, column=2, padx=5, pady=5, sticky='W')


Label___open_file_E = tk.Label(root, text="File E : ")
Label___open_file_E.grid(row=5, column=0, sticky='E')
Text___open_file_E = tk.Label(root, text="Please Select File...", height = 1, width = 45, bg="lightblue")
Text___open_file_E.grid(row=5, column=1, padx=5, pady=5, sticky='WE')
button___open_file_E  = tk.Button(root, text="Select...", command=lambda: open_file_button(Text___open_file_E))
button___open_file_E.grid(row=5, column=2, padx=5, pady=5, sticky='W')




# Text___open_file = tk.Text(root, height = 5, width = 52)
# Text___open_file.pack()

# # Data_Accumulate_Function
# button___Data_Accumulate = tk.Button(root, text="2. Data Store", command=Data_Accumulate_Button)
# button___Data_Accumulate.pack()

# label___Data_Accumulate = tk.Label(root, text="")
# label___Data_Accumulate.pack()

# # Data_Ploting_Function
# button___Data_Ploting_Function = tk.Button(root, text="3. Data Ploing", command=Data_Ploting_Function)
# button___Data_Ploting_Function.pack()

# label___Data_Ploting_Function = tk.Label(root, text="")
# label___Data_Ploting_Function.pack()

root.mainloop()
