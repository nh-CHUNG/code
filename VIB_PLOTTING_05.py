import tkinter as tk
from tkinter import filedialog



###### def open_file() Button:
def open_file():
    label_ploting_file['text'] = " "

    file_path = filedialog.askopenfilename(filetypes=[('Excel File','.xls'),('Excel File','.xlsx'),('Excel File','.csv'),('All File','.*')])
    file_path = file_path.replace('\\', '/')
    label_open_file['text'] = file_path
    # print(label_open_file['text'])



###### def ploting_file() Button:
def ploting_file():
    if label_open_file['text'] == 'Please Select File' :
        label_ploting_file['text'] = "Error : Please Select File"
    else :
        label_ploting_file['text'] = "Ploting Now"
        file_path  = label_open_file['text']
        ##### ploting
        ploting(file_path)



##### ploting
def ploting(input_path):
    import xlwings as xw
    import numpy as np
    import matplotlib
    import matplotlib.pyplot as plt
    import os
    # import pandas as pd
    ##### 불러온 파일 정리1
    xw.App(visible=False)
    # xw.App(visible=True)
    book = xw.Book(input_path)
    # book.sheets(1).range('1:1').delete()
    # book.sheets(1).range('A:A').insert("right")
    # book.sheets(1).range('A1').value = "Time"

    print(book.sheets(1).range('A1').current_region)
    print(book.sheets(1).range('A1').current_region.last_cell.row)
    print(book.sheets(1).range('A1').current_region.last_cell.column)

    row_size = book.sheets(1).range('A1').current_region.last_cell.row
    column_size = book.sheets(1).range('A1').current_region.last_cell.column

    ##### 불러온 파일 정리2
    np_input = book.sheets(1).used_range.options(np.array).value
    book.app.kill()

    ##### Re-Make data
    time_step = 0.001
    time = np.arange(0,(row_size-2)*time_step,time_step)
    x = np_input[1:row_size-1,1]
    y = np_input[1:row_size-1,2]
    z_1 = np_input[1:row_size-1,3]-1
    RMS = ((x**2 + y**2 + z_1**2)/3)**0.5
    del np_input

    x_max_time  = time[x.argmax()]
    x_max = x.max()
    x_min = x.min()

    y_max_time  = time[y.argmax()]
    y_max = y.max()
    y_min = y.min()

    z_1_max_time  = time[z_1.argmax()]
    z_1_max = z_1.max()
    z_1_min = z_1.min()

    RMS_max_time  = time[RMS.argmax()]
    RMS_max = RMS.max()
    RMS_min = RMS.min()


    ###### Plot
    fig, axs = plt.subplots(2,2 , figsize=(13, 5))

    matplotlib.rcParams['font.family'] ='Malgun Gothic' # 한글 폰트 깨짐 방지를 위한 코드
    matplotlib.rcParams['axes.unicode_minus'] = False # 한글 폰트 깨짐 방지를 위한 코드
    
    print_filename = os.path.basename(input_path)
    fig.suptitle(print_filename, fontsize=16)

    sizes = 0.50
    x_position_coefficient = time.max()*2/100
    y_position_coefficient = 0.85

    def text_position_height(min_value,max_value,coefficient):
        return min_value + (max_value - min_value) * coefficient

    axs[0,0].scatter(time, x, sizes, c="red")
    axs[0,0].grid(True)
    axs[0,0].set_title('X Vibration',loc='left')
    axs[0,0].set_xlabel('time [sec]')
    axs[0,0].set_ylabel('[G]')
    axs[0,0].scatter(x_max_time, x_max, sizes*50, c="red")
    text= "time={:.2f}\nx_max={:.2f}".format(x_max_time, x_max)
    axs[0,0].annotate(text, xy=( x_max_time + x_position_coefficient, text_position_height(x_min,x_max,y_position_coefficient) ) )


    axs[0,1].scatter(time, y, sizes, c="green")
    axs[0,1].grid(True)
    axs[0,1].set_title('Y Vibration',loc='left')
    axs[0,1].set_xlabel('time [sec]')
    axs[0,1].set_ylabel('[G]')
    axs[0,1].scatter(y_max_time, y_max, sizes*50, c="green")
    text= "time={:.2f}\ny_max={:.2f}".format(y_max_time, y_max)
    axs[0,1].annotate(text, xy=(y_max_time + x_position_coefficient, text_position_height(y_min,y_max,y_position_coefficient) ) )


    axs[1,0].scatter(time, z_1, sizes, c="blue")
    axs[1,0].grid(True)
    axs[1,0].set_title('Z-1 Vibration',loc='left')
    axs[1,0].set_xlabel('time [sec]')
    axs[1,0].set_ylabel('[G]')
    axs[1,0].scatter(z_1_max_time, z_1_max, sizes*50, c="blue")
    text= "time={:.2f}\nz_1_max={:.2f}".format(z_1_max_time, z_1_max)
    axs[1,0].annotate(text, xy=(z_1_max_time + x_position_coefficient, text_position_height(z_1_min,z_1_max,y_position_coefficient) ) )


    axs[1,1].scatter(time, RMS, sizes, c="orange")
    axs[1,1].grid(True)
    axs[1,1].set_title('RMS',loc='left')
    axs[1,1].set_xlabel('time [sec]')
    axs[1,1].set_ylabel('[G]')
    axs[1,1].scatter(RMS_max_time, RMS_max, sizes*50, c="orange")
    text= "time={:.2f}\nRMS_max={:.2f}".format(RMS_max_time, RMS_max)
    axs[1,1].annotate(text, xy=(RMS_max_time + x_position_coefficient, text_position_height(RMS_min,RMS_max,y_position_coefficient) ) )

    label_ploting_file['text'] = "Finish"

    fig.tight_layout()
    plt.show()





##################################################
##################################################
##################################################



##### 창 만들기
root = tk.Tk()
# width, height = 500, 25 # 창 크기 값 설정
# get_path = None
root.geometry("400x100") # 창 크기 설정
root.resizable(True, True) # 크기 조정 가능 여부
root.title('AVS Vibration Result Plotting') # 창 제목 설정

# open_file
button_open_file = tk.Button(root, text="1. Select File", command=open_file)
button_open_file.pack()

label_open_file = tk.Label(root, text="Please Select File")
label_open_file.pack()


# ploting_file
button_ploting_file = tk.Button(root, text="2. Ploting", command=ploting_file)
button_ploting_file.pack()

label_ploting_file = tk.Label(root, text="")
label_ploting_file.pack()







root.mainloop()
