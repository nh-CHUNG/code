import tkinter as tk
from tkinter import filedialog


###### def open_file() Button:
def open_file():
    global open_file_path
    label___Level_Column_Adding_Function['text'] = ""
    [open_file_path, open_file_name] = file_info()
    label___open_file['text'] = open_file_name


###### def file_info():
def file_info():
    # file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xlsx')])
    file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xls'),('Excel File','.xlsx'),('Excel File','.csv'),('All File','.*')])
    file_path = file_path.replace('\\', '/')

    index = file_path.rfind("/")
    file_name = file_path[index+1:len(file_path)]

    return file_path, file_name


###### def Level_Column_Adding_file() Button:
def Level_Column_Adding_file():
    if label___open_file['text'] == 'Please Select File' :
        label___Level_Column_Adding_Function['text'] = "Error : Please Select File"
        # print(label___open_file['text'])
    else :
        label___Level_Column_Adding_Function['text'] = "Leveling Now"
        open_file_path
        # print(label___open_file['text'])
        ##### ploting
        Level_Column_Adding_Function(open_file_path)
        label___Level_Column_Adding_Function['text'] = "Finish"


###### def Level_Column_Adding_Function(input_path):
def Level_Column_Adding_Function(input_path):
    import xlwings as xw
    import numpy as np

    ##### 파일 불러오기
    # input_path = r'C:\Users\namho.chung\Desktop\STUDY\BOM LEVELING\BOMReport_20250714074652.xlsx' # r'brabra' Raw String 


    ##### 불러온 파일 정리1
    xw.App(visible=False)
    book = xw.Book(input_path)
    row_size = book.sheets(1).range('A1').current_region.last_cell.row
    column_size = book.sheets(1).range('A1').current_region.last_cell.column
    np_excel_data = book.sheets(1).used_range.options(np.array).value


    ##### 불러온 파일 정리2 - Level 항목 추출 및 정리
    Level_arry = np_excel_data[1:,0]
    del np_excel_data
    Level_arry = Level_arry.astype(float) # 문자를 float으로 
    Level_arry = Level_arry.astype(int) # float를 int로
    Level_max_value = int(max(Level_arry)) # 최대값 확인


    ##### Leveling 열 추가를 위한 작업
    i = 0
    for i in range(Level_max_value):
        temp_arry = np.zeros_like(Level_arry) # temp_arry 초기화
        Current_Level = Level_max_value - i
        
        ### Level 열 추가
        book.sheets(1).range('B:B').insert("right")
        book.sheets(1).range('B:B').column_width = 5
        book.sheets(1).range('B1').value = Current_Level

        ### 추가 된 열에 Level 입력을 위한 사전 작업
        # input_Level_arry = book.sheets(1).range((2,2),(row_size,2)).value
        bool_index = Level_arry <= Current_Level # 비교 연산
        temp_arry = bool_index * Level_arry
        temp_arry = np.where(temp_arry == 0, 1510, temp_arry) # replace를 위한 임의 숫자를 삽입
        temp_arry = temp_arry.reshape(-1, 1) # 가로 데이터를 세로 데이터로 변환

        ### 추가 된 열에 Level 입력 및 정리
        book.sheets(1).range('B2').value = temp_arry # 정리된 Level 입력
        # print(' ///// EXE : Current Level ///// ')    
        book.sheets(1).used_range.api.Replace("1510", "") # 입력된 값 중 불필요 값 정리
        # print(' ///// EXE : Replace ///// ')    
        
 
    ##### 작업 완료 파일 저장
    index = input_path.rfind(".")
    input_path = input_path[0:index] + "_LEVEL.xlsx"
    book.save(input_path)
    book.app.kill()


##### 창 만들기 #####
root = tk.Tk()


# width, height = 500, 25 # 창 크기 값 설정
# get_path = None
root.geometry("400x100") # 창 크기 설정
root.resizable(True, True) # 크기 조정 가능 여부
root.title('BOMREport Leveling Column Adding') # 창 제목 설정

# open_file 
button___open_file = tk.Button(root, text="1. Select File", command=open_file)
button___open_file.pack()

label___open_file = tk.Label(root, text="Please Select File")
label___open_file.pack()


# Level_Column_Adding_Function
button___Level_Column_Adding_Function = tk.Button(root, text="2. Level Column Adding", command=Level_Column_Adding_file)
button___Level_Column_Adding_Function.pack()

label___Level_Column_Adding_Function = tk.Label(root, text="")
label___Level_Column_Adding_Function.pack()

root.mainloop()
