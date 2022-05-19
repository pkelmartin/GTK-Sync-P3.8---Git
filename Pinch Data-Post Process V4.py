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

                    if files.startswith('230YT'):
                        side = 'Left'
                    if files.startswith('230YV'):
                        side = 'Right'

                    if position.lower() == 'user posit':
                        position = 'User Position'
                    if position.lower() == 'maximum de':
                        position = 'Maximum Deflection'

                    df1 = pd.read_excel(test_path)


                    ADC_CLOCK = df1[df1.columns[0]]
                    Fore_ADC = df1[df1.columns[1]]
                    Force_clock = df1[df1.columns[2]]
                    force_list = df1[df1.columns[3]]
                    Z_Posi = df1[df1.columns[4]]



                    df2 = pd.DataFrame()
                    df2.insert(0, 'Force Clock', Force_clock)
                    df2.insert(1, 'Force', force_list)
                    df2.insert(2, 'Z_Position', Z_Posi)

                    df3 = pd.DataFrame()
                    df3.insert(0, 'ADC Clock', ADC_CLOCK)
                    df3.insert(1, 'Pinch ADC', Fore_ADC)


                    i = 0
                    for z in df2['Z_Position']:
                        if i > 100:
                            if z < 0.001:
                                break
                        i = i + 1

                    first_press = df2.iloc[:i]
                    two_start = df2.iloc[i]
                    two_three_four = df2[i:]
                    Second_press_time = two_start['Force Clock']

                    i = 0
                    for z in two_three_four['Z_Position']:
                        if i > 100:
                            if z < 0.001:
                                break
                        i = i + 1

                    second_press = two_three_four[:i]
                    three_four = two_three_four[i:]
                    three_start = two_three_four.iloc[i]
                    Third_press_time = three_start['Force Clock']

                    i = 0
                    for z in three_four['Z_Position']:
                        if i > 100:
                            if z < 0.001:
                                break
                        i = i + 1

                    third_press = three_four[:i]
                    fourth_press = three_four[i:]
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
                    for press in press_frames_force:

                        adc_frame = press_frames_adc[l]

                        force_1 = press['Force']

                        index = press.index

                        count = 0
                        for j in index:
                            if count == 0:
                                k = j
                                break
                            count = count + 1


                        for force in force_1:
                            if force_1[k] < force_1[k+1] and force_1[k+1] < force_1[k+2] and force_1[k+2] < force_1[k+3] and force_1[k+3] < force_1[k+4] and force_1[k+4] < force_1[k+5] and force_1[k+5] < force_1[k+6] and force_1[k+6] < force_1[k+7]:
                                break
                            k = k + 1

                        force_start_row = press.loc[k]


                        adc_5 = adc_frame['Pinch ADC']

                        index1 = adc_frame.index

                        count1 = 0
                        for j1 in index1:
                            if count == 0:
                                n = j1
                                break
                            count1 = count1 + 1


                        for adc in adc_5:
                            if adc_5[n] < adc_5[n+1] and adc_5[n+1] < adc_5[n+2] and adc_5[n+2] < adc_5[n+3] and adc_5[n+3] < adc_5[n+4] and adc_5[n+4] < adc_5[n+5] and adc_5[n+5] < adc_5[n+6] and adc_5[n+6] < adc_5[n+7]:
                                break
                            n = n + 1

                        adc_start_row = adc_frame.loc[n]

                        start_distance = force_start_row['Z_Position']

                        max_distance = max(press['Z_Position'])

                        Max_Displacement = max_distance - start_distance

                        Force_Start_Time = force_start_row['Force Clock']

                        start_Force_ADC_1 = adc_frame.loc[adc_frame['ADC Clock'] >= Force_Start_Time]
                        start_Force_ADC_2 = start_Force_ADC_1.iloc[0]

                        START_ADC_TIME = start_Force_ADC_2['ADC Clock']

                        START_ADC_DISTANCE_ROW_1 = press.loc[press['Force Clock'] >= START_ADC_TIME]
                        START_ADC_DISTANCE_ROW_2 = START_ADC_DISTANCE_ROW_1.iloc[0]

                        START_ADC_DISTANCE = START_ADC_DISTANCE_ROW_2['Z_Position']
                        START_ADC_FORCE = START_ADC_DISTANCE_ROW_2['Force']

                        Travel_to_ADC = START_ADC_DISTANCE - start_distance


                        oo = j
                        pp = 0
                        found = 0
                        last_force = 0
                        force_list_1 = []
                        difference_l = []
                        for force in force_1:
                            force_list_1.append(force)
                            max_force_1 = max(force_list_1)
                            f_d = force - last_force
                            difference_l.append(f_d)
                            last_force = force
                            if max_force_1 > 4 and force < 0.01:
                                found = 1
                                break
                            oo = oo + 1
                            pp = pp + 1


                        end_force_row = press.loc[oo]


                        end_force_time = end_force_row['Force Clock']

                        end_Force_ADC_1 = adc_frame.loc[adc_frame['ADC Clock'] >= end_force_time]
                        end_Force_ADC_2 = end_Force_ADC_1.iloc[0]

                        end_force_ADC_V = end_Force_ADC_2['Pinch ADC']

                        min_start_adc_df = df1.iloc[:k]
                        min_start_adc = min(min_start_adc_df['Pinch ADC'])

                        RTZ_AT_ZERO_N = end_force_ADC_V - min_start_adc

                        max_adc = max(adc_frame['Pinch ADC'])

                        RTZ_PCT = (RTZ_AT_ZERO_N/(max_adc-min_start_adc))

                        if not os.path.exists(full_path+'/'+folders+' Plots'):
                            os.makedirs(full_path+'/'+folders+' Plots')
                        if not os.path.exists(full_path+'/'+folders+' Plots/'+sn+'-'+position+' Press '+str(l+1)+'.png'):

                            fig1 = plt.figure()
                            fig1, ax1 = plt.subplots()

                            adctime = adc_frame['ADC Clock']

                            adc_list = adc_frame['Pinch ADC']

                            adc_normal = []
                            for adc_value in adc_list:
                                # adc_1000 = (adc_value/1000)-min_start_adc/1000
                                adc_normal.append(adc_value)

                            force_list_p = press['Force']
                            force_clock = press['Force Clock']
                            z_pos_list = press['Z_Position']

                            ax1.plot(force_clock,force_list_p, label='Load Cell')
                            # ax1.plot(adctime,adc_normal, label='Normalised ADC')
                            ax1.plot(force_clock,z_pos_list, label='Z Position')
                            # plt.plot(end_force_time, , '*', label='Force Start')

                            ax1.axvline(x=(Force_Start_Time),color='b',label='Start Force', linestyle='--')
                            ax1.axvline(x=(end_force_time),color='r',label='End Force', linestyle='--')
                            ax2 = ax1.twinx()
                            ax2.plot(adctime,adc_normal, label='Pinch ADC', color='green')


                            ax1.legend(title='Force',  prop={'size': 6}, loc=2)
                            ax2.legend(title='ADC',  prop={'size': 6}, loc=1)





                            plt.savefig(full_path+'/'+folders+' Plots/'+sn+'-'+position+' Press '+str(l+1)+'.png',dpi=500)
                            plt.close()

                        Data_Values = {'S/N': [sn],
                                       'Side': [side],
                                       'Position': [position],
                                       'Travel to ADC': [Travel_to_ADC],
                                       '7N Travel': [Max_Displacement],
                                       'RTZ @0N': [RTZ_AT_ZERO_N],
                                       'RTZ %': [RTZ_PCT],
                                       '0N ADC': [min_start_adc],
                                       '6N ADC': [max_adc],
                                       'ADC Range': [max_adc-min_start_adc],
                                       'Force to ADC': [START_ADC_FORCE]
                                       }

                        df32 = pd.DataFrame(Data_Values)

                        # df33 = df33.append(df32)

                        df33 = pd.concat([df33, df32])    #append going away



                        l = l + 1

                        print('debug')
    df33.to_excel(full_path+'/'+'GTK ' + folders +' results.xlsx', index=False)
print('Done!')