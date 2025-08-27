import tkinter as tk
from tkinter import filedialog


###### def open_file() Button:
def open_file():
    global open_file_path
    label___Data_Re_Making_Function['text'] = ""
    [open_file_path, open_file_name] = file_info()
    label___open_file['text'] = open_file_name
    
    # print(label___open_file['text'])

###### def file_info():
def file_info():
    # file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xlsx')])
    file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xls'),('Excel File','.xlsx'),('Excel File','.csv'),('All File','.*')])
    file_path = file_path.replace('\\', '/')

    index = file_path.rfind("/")
    file_name = file_path[index+1:len(file_path)]

    return file_path, file_name


###### def Data_Re_Making_Bottom() Button:
def Data_Re_Making_Bottom():
    if label___open_file['text'] == 'Please Select File' :
        label___Data_Re_Making_Function['text'] = "Error : Please Select File"
        # print(label___open_file['text'])
    else :
        label___Data_Re_Making_Function['text'] = "Re-Making Now"
        open_file_path
        # print(label___open_file['text'])
        ##### ploting
        Data_Re_Making_Function(open_file_path)
        label___Data_Re_Making_Function['text'] = "Finish"


###### def Data_Re_Making_Function(input_path):
def Data_Re_Making_Function(input_path):
    import xlwings as xw
    import numpy as np

    ##### 파일 불러오기
    # input_path = r'C:\Users\namho.chung\Desktop\STUDY\BOM LEVELING\BOMReport_20250714074652.xlsx' # r'brabra' Raw String 


    ##### 불러온 파일 정리1
    # xw.App(visible = False)
    xw.App(visible = True)
    book = xw.Book(input_path)
    row_size = book.sheets(1).range('A1').current_region.last_cell.row
    column_size = book.sheets(1).range('A1').current_region.last_cell.column

    ##### 불러온 파일 정리2
    # 첫 행에 불필요한 정보 삭제 및 수정 
    # book.sheets(1).range("1:1").api.Replace("(자동화팀)", "[자동화팀]")  
    # book.sheets(1).range("1:1").api.Replace("(내부)", "[내부]")  
    # book.sheets(1).range("1:1").api.Replace("(외부)", "[외부]")  
    # book.sheets(1).range("1:1").api.Replace("사용자정의 필드 (", "")  
    # book.sheets(1).range("1:1").api.Replace(")", "")  

    np_excel_data = book.sheets(1).used_range.options(np.array).value
    del np_excel_data

    ##### 불러온 파일 정리3
    # 상태 열에 순서 번호 입히기
    # found_cell = book.sheets(1).used_range.api.Find(What="상태")
    # Cell_address = found_cell._inner.Address
    # Cell_address = Cell_address.replace("$","")
    # Cell_address = Cell_address[0]

    Cell_address = find_cell_info(book,"상태")

    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("발굴", "01. 발굴")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("Concept설계", "02. Concept설계")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("상세 설계", "03. 상세 설계")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("제작", "04. 제작")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("Sample Test", "05. Sample Test")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("Revision", "06. Revision")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("Field Test", "07. Field Test")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("COA", "08. COA")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("PCCB", "09. PCCB")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("완료", "10. 완료")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("Drop", "11. Drop")

    ##### 불러온 파일 정리4
    # 비용±(내부) 중 N/A는 빈칸으로 변경
    # found_cell = book.sheets(1).used_range.api.Find(What="비용±(내부)")
    # Cell_address = found_cell._inner.Address
    # Cell_address = Cell_address.replace("$","")
    # Cell_address = Cell_address[0]

    Cell_address = find_cell_info(book, "비용 변화(내부)")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("N/A", "")

    ##### 불러온 파일 정리5
    # 비용±(외부) 중 N/A는 빈칸으로 변경
    # found_cell = book.sheets(1).used_range.api.Find(What="비용±(외부)")
    # Cell_address = found_cell._inner.Address
    # Cell_address = Cell_address.replace("$","")
    # Cell_address = Cell_address[0]

    Cell_address = find_cell_info(book, "비용 변화(외부)")
    book.sheets(1).range(f'{Cell_address}:{Cell_address}').api.Replace("N/A", "")


    ##### 작업 완료 파일 저장
    index = input_path.rfind(".")
    input_path = input_path[0:index] + "_Re_Making.xlsx"
    book.save(input_path)
    book.app.kill()


###### def find_cell_info():
def find_cell_info(input_book, target_text):
    found_cell = input_book.sheets(1).used_range.api.Find(What=target_text)
    found_address = found_cell._inner.Address

    index_front = found_address.find("$")
    index_rear = found_address.rfind("$")

    found_address = found_address[index_front+1:index_rear]

    return found_address



##### 창 만들기 #####
root = tk.Tk()
# width, height = 500, 25 # 창 크기 값 설정
# get_path = None
root.geometry("400x100") # 창 크기 설정
root.resizable(True, True) # 크기 조정 가능 여부
root.title('JIRA Filter Exported Data Re-Building') # 창 제목 설정

# open_file 
button___open_file = tk.Button(root, text="1. Select File", command=open_file)
button___open_file.pack()

label___open_file = tk.Label(root, text="Please Select File")
label___open_file.pack()

# Data_Re_Making_Function
button___Data_Re_Making_Function = tk.Button(root, text="2. Data Re-Making", command=Data_Re_Making_Bottom)
button___Data_Re_Making_Function.pack()

label___Data_Re_Making_Function = tk.Label(root, text="")
label___Data_Re_Making_Function.pack()

root.mainloop()
