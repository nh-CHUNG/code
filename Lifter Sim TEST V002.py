import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import matplotlib.cm as cm
import matplotlib.colors as colors
import numpy as np

def Lift_test():
    ### Initial Data ###
    ###### Time
    TIME = 0 ; TIME_End = 10 ; TIME_step = 1 # Unit : se
    
    ###### Left[Floor][Destination] = value = waiting time
    FLOOR = 4 
    Divider = 10
    Left = np.eye( FLOOR ) 
    i=0 ; j=0
    for i in range(FLOOR):
        for j in range(FLOOR):
            if Left[i][j]==1:
                Left[i][j] = 0
            else :
                Left[i][j] = Wait_time = int(np.random.rand()*10)
    print("Left[][] = ",Left)

    ###### Lift P V A
    Lift_Position = 0
    Lift_Velocity = 3 # meter/sec
    Lift_Acceleration = 1 # meter/sec^2

    ###### Plot 설정
    plt.ion() 

    subplot_mosaic = np.zeros((4, 2), dtype=int)
    i=0
    for i in range(FLOOR):
        subplot_mosaic[i][0] = FLOOR - i
    print("mosaic[][]= ",subplot_mosaic)
    axs = plt.figure(layout="constrained").subplot_mosaic(subplot_mosaic, empty_sentinel=0)
    
    axis_Floor = [f for f in range(0, FLOOR)]
    
    ######### Plot 색상 설정
    cmap = cm.Pastel2
    norm = colors.Normalize(vmin=min(axis_Floor), vmax=max(axis_Floor))
    bar_colors = [cmap(norm(value)) for value in axis_Floor]


    ######### 초기 Plot 생성
    i = 0
    y_max = Divider
    
    for i in range(0,4):
        axs[i+1].bar(axis_Floor, Left[i], width=1, edgecolor="white", linewidth=0.7, color=bar_colors)        
        axs[i+1].set_title(f"{i} Floor",loc='right', size=8)
        axs[i+1].set_xticks([xticks for xticks in range(min(axis_Floor), max(axis_Floor)+1)])
        if i==0:
            axs[i+1].set_xlabel('Destination [floor]', size=6)
        axs[i+1].set_yticks([yticks for yticks in range(0, y_max+1, 3)])
        axs[i+1].set_ylabel('Time Sent [sec]', size=6)

    
    # while TIME <= TIME_End:
    #     try:
    #         print(Left)
    #         print("TIME =",TIME)

    #         # 반송할 대상 선정
    #         max_adress = np.argmax(Left)
    #         Floor = max_adress // 4 ; Destination = max_adress % 4
    #         Left[max_adress // 4][max_adress % 4]

    #         axs[0,0].grid(True)
    #         axs[0,0].set_title('X Vibration',loc='left')
    #         axs[0,0].set_xlabel('time [sec]')
    #         axs[0,0].set_ylabel('[G]')
    #         axs[0,0].scatter(1, Left[0,][0], 50, c="red")
    #         axs.plot(x_data, y_data)


    #         TIME = TIME + TIME_step
    #     except KeyboardInterrupt:
    #         # Ctrl+C를 누르면 루프 종료
    #         break

    plt.ioff() # 대화형 모드 비활성화
    plt.tight_layout()
    plt.show()

    # Temp
    max_adress = np.argmax(Left)
    Floor = max_adress // 4
    Destination = max_adress % 4
    print("Left [Floor(",Floor,")][Destination(",Destination,")] =",Left[max_adress // 4][max_adress % 4])

    print("STOP")


if __name__ == '__main__':
    Lift_test()
