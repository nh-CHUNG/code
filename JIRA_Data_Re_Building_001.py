import tkinter as tk
from tkinter import filedialog


###### def open_file() Button:
def open_file():
    label___Data_Re_Making_Function['text'] = ""

    # file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xlsx')])
    file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xls'),('Excel File','.xlsx'),('Excel File','.csv'),('All File','.*')])
    file_path = file_path.replace('\\', '/')
    label___open_file['text'] = file_path
    
    # print(label___open_file['text'])


###### def Data_Re_Making_Bottom() Button:
def Data_Re_Making_Bottom():
    if label___open_file['text'] == 'Please Select File' :
        label___Data_Re_Making_Function['text'] = "Error : Please Select File"
        # print(label___open_file['text'])
    else :
        label___Data_Re_Making_Function['text'] = "Re-Making Now"
        file_path  = label___open_file['text']
        # print(label___open_file['text'])
        ##### ploting
        Data_Re_Making_Function(file_path)
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
    book.sheets(1).range("1:1").api.Replace("(자동화팀)", "[자동화팀]")  
    book.sheets(1).range("1:1").api.Replace("(내부)", "[내부]")  
    book.sheets(1).range("1:1").api.Replace("(외부)", "[외부]")  
    book.sheets(1).range("1:1").api.Replace("사용자정의 필드 (", "")  
    book.sheets(1).range("1:1").api.Replace(")", "")  

    ##### 불러온 파일 정리3
    # 상태 열에 순서 번호 입히기
    found_cell = book.sheets(1).used_range.api.Find(What="상태")
    Status_address = found_cell._inner.Address
    Status_address = Status_address.replace("$","")
    Status_address = Status_address[0]
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("발굴", "1. 발굴")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("Concept설계", "2. Concept설계")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("상세 설계", "3. 상세 설계")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("제작", "4. 제작")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("Sample Test", "5. Sample Test")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("Revision", "6. Revision")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("Field Test", "7. Field Test")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("COA", "8. COA")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("PCCB", "9. PCCB")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("완료", "10. 완료")
    book.sheets(1).range(f'{Status_address}:{Status_address}').api.Replace("Drop", "11. Drop")


    np_excel_data = book.sheets(1).used_range.options(np.array).value
    del np_excel_data

               
    ##### 작업 완료 파일 저장
    index = input_path.rfind(".")
    input_path = input_path[0:index] + "_Re_Making.xlsx"
    book.save(input_path)
    book.app.kill()


##### 창 만들기 #####
root = tk.Tk()
# width, height = 500, 25 # 창 크기 값 설정
# get_path = None
root.geometry("400x100") # 창 크기 설정
root.resizable(True, True) # 크기 조정 가능 여부
root.title('JIRA exported Data Re-Making') # 창 제목 설정

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
