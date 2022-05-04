from tkinter import filedialog
import pandas as pd
import os
import tkinter.filedialog
from os import walk
import json
import math
import matplotlib.pyplot as plt

Test_Number = 1

# df33 = pd.DataFrame()

# File dialog to select folder containing excel files to import
full_path = tkinter.filedialog.askdirectory(initialdir='.')
os.chdir(full_path)

f1 = []
f2 = []
for (dirpath, dirnames, filenames) in walk(full_path):
    f1.extend(filenames)
    f2.extend(dirnames)

for folders in f2:
    df33 = pd.DataFrame()
    if folders != 'Plots':
        f3 = []
        for (dirpath2, dirnames2, filenames2) in walk(full_path+'/'+folders):
            f3.extend(filenames2)

        for files in f3:
            if files.startswith('230'):
                if files.endswith(".xlsx"):
                    test_path = full_path+'/'+folders+'/'+files
                    print(test_path)

                    sn = files[:14]
                    position = files[15:25]

                    df1 = pd.read_excel(test_path)
                    if files.startswith('230YT'):
                        side = 'Left'
                    if files.startswith('230YV'):
                        side = 'Right'


                    ADC_CLOCK = df1[df1.columns[0]]
                    Fore_ADC = df1[df1.columns[1]]
                    Force_clock = df1[df1.columns[2]]
                    force_list = df1[df1.columns[3]]
                    Z_Posi = df1[df1.columns[4]]



                    df2 = pd.DataFrame()
                    df2.insert(0, 'Force Clock', Force_clock)
                    df2.insert(1, 'Force', force_list)
                    df2.insert(2, 'Z Position', Z_Posi)

                    df3 = pd.DataFrame()
                    df3.insert(0, 'ADC Clock', ADC_CLOCK)
                    df3.insert(1, 'Grip ADC', Fore_ADC)

                    i = 0
                    for z in df2['Z Position']:
                        if i > 100:
                            if z < 0.001:
                                break
                        i = i + 1

                    first_press = df2.iloc[:i]
                    two_start = df2.iloc[i]
                    two_three_four = df2[i:]
                    Second_press_time = two_start['Force Clock']

                    i = 0
                    for z in two_three_four['Z Position']:
                        if i > 100:
                            if z < 0.001:
                                break
                        i = i + 1

                    second_press = two_three_four[:i]
                    three_four = two_three_four[i:]
                    three_start = two_three_four.iloc[i]
                    Third_press_time = three_start['Force Clock']

                    i = 0
                    for z in three_four['Z Position']:
                        if i > 100:
                            if z < 0.001:
                                break
                        i = i + 1

                    third_press = three_four[:i]
                    fourth_press = three_four[i:]
                    # four_star1 = three_four[i]
                    four_start = three_four.iloc[i]
                    Forth_press_time = four_start['Force Clock']

                    first_press_adc = df3.loc[df3['ADC Clock'] < Second_press_time]
                    Second_press_adc_1 = df3.loc[df3['ADC Clock'] >= Second_press_time]
                    Second_press_adc_2 = Second_press_adc_1.loc[Second_press_adc_1['ADC Clock'] < Third_press_time]
                    Third_press_adc_1 = df3.loc[df3['ADC Clock'] >= Third_press_time]
                    Third_press_adc_2 = Third_press_adc_1.loc[Third_press_adc_1['ADC Clock'] < Forth_press_time]
                    Fourth_press_adc = df3.loc[df3['ADC Clock'] >= Forth_press_time]

                    press_frames_force = [first_press, second_press, third_press, fourth_press]
                    press_frames_adc = [first_press_adc, Second_press_adc_2, Third_press_adc_2, Fourth_press_adc]

                    l = 0
                    for press in press_frames_force and press_frames_force:


                        adc_frame = press_frames_adc[l]

                        print('debug')
                        max_ADC = max(adc_frame['Grip ADC'])

                        min_ADC = min(adc_frame['Grip ADC'])

                        ADC_Delta = max_ADC - min_ADC

                        force_list = []
                        for force in press['Force']:
                            force_list.append(force)
                            if force > 0.2:
                                break

                        force_list.reverse()

                        index = press.index

                        count = 0
                        for j in index:
                            if count == 0:
                                start = j
                                break
                            count = count + 1

                        # j = index[-1]

                        n = 0
                        last_value = force_list[0]
                        for force_2 in force_list:
                            n = n + 1
                            if force_2 < 0:
                                break
                            last_value = force_2


                        len_fl = len(force_list)

                        force_start_row = press.loc[start+len_fl-n]

                        force_start_time = force_start_row['Force Clock']

                        force_start_dist = force_start_row['Z Position']

                        max_travel = max(press['Z Position'])

                        Z_delta = max_travel-force_start_dist

                        force_start_adc_row = adc_frame.loc[adc_frame['ADC Clock'] > force_start_time]

                        if not os.path.exists(full_path+'/'+folders+' Plots'):
                            os.makedirs(full_path+'/'+folders+' Plots')
                        if not os.path.exists(full_path+'/'+folders+' Plots/'+sn+'-'+position+' Press '+str(l+1)+'.png'):

                            fig1 = plt.figure()

                            adctime = adc_frame['ADC Clock']

                            adc_list = adc_frame['Grip ADC']

                            adc_normal = []
                            for adc_value in adc_list:
                                adc_1000 = (adc_value/1000)-max_ADC/1000
                                adc_normal.append(-1*adc_1000)

                            force_list_p = press['Force']
                            force_clock = press['Force Clock']
                            z_pos_list = press['Z Position']

                            plt.plot(force_clock,force_list_p, label='Load Cell')
                            plt.plot(adctime,adc_normal, label='Normalised ADC')
                            plt.plot(force_clock,z_pos_list, label='Z Position')
                            # plt.plot(end_force_time, , '*', label='Force Start')

                            plt.axvline(x=(force_start_time),color='b',label='Start Force', linestyle='--')
                            # plt.axvline(x=(end_force_time),color='g',label='End Force', linestyle='--')


                            plt.legend(title='KPI Sources')





                            plt.savefig(full_path+'/'+folders+' Plots/'+sn+'-'+position+' Press '+str(l+1)+'.png',dpi=500)
                            plt.close()



                        Data_Values = {'S/N': [sn],
                                       'Side': [side],
                                       'Position': [position],
                                       'ADC Delta': [ADC_Delta],
                                       'Pressed ADC': [min_ADC],
                                       'UnPressed ADC': [max_ADC],
                                       'Max Displacement': [Z_delta]
                                       }

                        df32 = pd.DataFrame(Data_Values)

                        # df33 = df33.append(df32)

                        df33 = pd.concat([df33, df32])    #append going away

                        l = l + 1

    df33.to_excel(full_path+'/'+'GTK ' + folders +' results.xlsx', index=False)
print('debug')
