import pandas as pd
import os
from os import walk
import matplotlib.pyplot as plt


def grip_trigger_post_processing_v4(full_path):

    Test_Number = 1

    # df33 = pd.DataFrame()

    # File dialog to select folder containing excel files to import
    # full_path = tkinter.filedialog.askdirectory(initialdir='.')
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
                        Grip_ADC = df1[df1.columns[1]]
                        Force_clock = df1[df1.columns[2]]
                        force_list = df1[df1.columns[3]]
                        Z_Posi = df1[df1.columns[4]]

                        df2 = pd.DataFrame()
                        df2.insert(0, 'Force Clock', Force_clock)
                        df2.insert(1, 'Force', force_list)
                        df2.insert(2, 'Z_Position', Z_Posi)

                        df3 = pd.DataFrame()
                        df3.insert(0, 'ADC Clock', ADC_CLOCK)
                        df3.insert(1, 'Grip ADC', Grip_ADC)

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
                        for press in press_frames_force and press_frames_force:

                            adc_frame = press_frames_adc[l]

                            previous_force = 0
                            force_list_d = []
                            force_list_t = []
                            count = 0
                            force_list_10 = []
                            for force_3 in press['Force']:
                                force_4 = force_3 - previous_force
                                force_list_t.append(force_3)
                                force_list_d.append(force_4)
                                previous_force = force_3
                                if count > 200:
                                    if force_list_t[count-10] > .5:
                                        ten_back = force_3 - force_list_t[count-10]
                                        force_list_10.append(ten_back)
                                        if ten_back > 0.1:
                                            break
                                count = count + 1

                            hardstop_force_row_1 = press.iloc[count-1]

                            force_zone = force_list_t[(count-10):count]

                            force_zone_d = [(force_zone[1]-force_zone[0]), (force_zone[2]-force_zone[1]), (force_zone[3]-force_zone[2]), (force_zone[4]-force_zone[3]),(force_zone[5]-force_zone[4]), (force_zone[6]-force_zone[5]), (force_zone[7]-force_zone[6]), (force_zone[8]-force_zone[7]), (force_zone[9]-force_zone[8])]

                            force_zone_d_2 = [(force_zone[1]-force_zone[0]), (force_zone[2]-force_zone[0]), (force_zone[3]-force_zone[0]), (force_zone[4]-force_zone[0]),(force_zone[5]-force_zone[0]), (force_zone[6]-force_zone[0]), (force_zone[7]-force_zone[0]), (force_zone[8]-force_zone[0]), (force_zone[9]-force_zone[0])]

                            count_2 = 0
                            for force_6 in force_zone_d_2:
                                if force_6 > 0.005:
                                    offset_count = count_2
                                    break
                                count_2 = count_2 + 1

                            hardstop_force_row = press.iloc[count-(10-count_2)]

                            hardstop_force_row_time = hardstop_force_row['Force Clock']

                            adc_hard_stop_1 = adc_frame.loc[adc_frame['ADC Clock'] <= hardstop_force_row_time]
                            adc_hard_stop_2 = adc_hard_stop_1.iloc[-1]

                            reverse = []
                            for force_7 in press['Force']:
                                reverse.append(force_7)

                            reverse.reverse()

                            previous_force = 0
                            force_list_d = []
                            force_list_t = []
                            count = 0
                            force_list_10 = []
                            for force_3 in reverse:
                                force_4 = force_3 - previous_force
                                force_list_t.append(force_3)
                                force_list_d.append(force_4)
                                previous_force = force_3
                                if count > 200:
                                    if force_list_t[count-10] > .5:
                                        ten_back = force_3 - force_list_t[count-10]
                                        force_list_10.append(ten_back)
                                        if ten_back > 0.1:
                                            break
                                count = count + 1

                            hardstop_force_row_3 = press.iloc[len(reverse) - count]

                            force_zone = force_list_t[(count-10):count]

                            force_zone_d = [(force_zone[1]-force_zone[0]), (force_zone[2]-force_zone[1]), (force_zone[3]-force_zone[2]), (force_zone[4]-force_zone[3]),(force_zone[5]-force_zone[4]), (force_zone[6]-force_zone[5]), (force_zone[7]-force_zone[6]), (force_zone[8]-force_zone[7]), (force_zone[9]-force_zone[8])]

                            force_zone_d_2 = [(force_zone[1]-force_zone[0]), (force_zone[2]-force_zone[0]), (force_zone[3]-force_zone[0]), (force_zone[4]-force_zone[0]),(force_zone[5]-force_zone[0]), (force_zone[6]-force_zone[0]), (force_zone[7]-force_zone[0]), (force_zone[8]-force_zone[0]), (force_zone[9]-force_zone[0])]

                            count_2 = 0
                            for force_6 in force_zone_d_2:
                                if force_6 > 0.005:
                                    offset_count = count_2
                                    break
                                count_2 = count_2 + 1



                            hardstop_force_row_2 = press.iloc[len(reverse) - count+12-count_2]

                            hardstop_force_row_2_time = hardstop_force_row_2['Force Clock']

                            adc_hard_stop_3 = adc_frame.loc[adc_frame['ADC Clock'] >= hardstop_force_row_2_time]
                            adc_hard_stop_4 = adc_hard_stop_3.iloc[0]




                            print('debug')
                            max_ADC = max(adc_frame['Grip ADC'])

                            min_ADC = min(adc_frame['Grip ADC'])

                            ADC_Delta = max_ADC - min_ADC

                            k = 0
                            force_list = []
                            for force in press['Force']:
                                force_list.append(force)
                                if force > 0.25:
                                    break
                                k = k + 1

                            force_list.reverse()

                            j = 0
                            last_value = force_list[0]
                            for force_2 in force_list:
                                if force_2 < 0.001:
                                    break
                                last_value = force_2
                                j = j + 1

                            force_start_row = press.iloc[k-j]

                            force_distance_start = force_start_row['Z_Position']

                            force_start_time = force_start_row['Force Clock']

                            force_start_force = force_start_row['Force']

                            max_distance = max(press['Z_Position'])

                            Travel = max_distance - force_distance_start

                            if not os.path.exists(full_path+'/'+folders+' Plots'):
                                os.makedirs(full_path+'/'+folders+' Plots')
                            if not os.path.exists(full_path+'/'+folders+' Plots/'+sn+'-'+position+' Press '+str(l+1)+'.png'):


                                force_list.reverse()
                                fig1, ax1 = plt.subplots()
                                plt.title(sn+'-'+position+' Press '+str(l+1))

                                adctime = adc_frame['ADC Clock']

                                adc_list = adc_frame['Grip ADC']

                                adc_normal = []
                                for adc_value in adc_list:
                                    # adc_1000 = (adc_value/3000)-max_ADC/3000
                                    adc_normal.append(1*adc_value)

                                force_list_p = press['Force']
                                force_clock = press['Force Clock']
                                z_pos_list = press['Z_Position']

                                ax1.plot(force_clock,force_list_p, label='Load Cell')
                                ax1.plot(force_clock,z_pos_list, label='Z Position')
                                ax1.plot(force_start_time, force_start_force, '*', label='Force Start')

                                ax1.axvline(hardstop_force_row['Force Clock'], label='Hard Stop 1', color='red')

                                ax1.axvline(hardstop_force_row_2['Force Clock'], label='Hard Stop 2', color='red')
                                ax2 = ax1.twinx()
                                ax2.plot(adctime,adc_normal, label='Grip ADC', color='green')
                                ax2.axhline(y=700, color='purple', label='700 ADC Line')

                                ax1.legend(title='Force',  prop={'size': 6}, loc=2)
                                ax2.legend(title='ADC',  prop={'size': 6}, loc=1)

                                plt.savefig(full_path+'/'+folders+' Plots/'+sn+'-'+position+' Press '+str(l+1)+'.png',dpi=500)
                                plt.close()


                            if adc_hard_stop_2['Grip ADC'] - min_ADC < 25:
                                Saturated = 'Yes'
                            else:
                                Saturated = 'No'

                            Data_Values = {'S/N': [sn],
                                           'Side': [side],
                                           'Position': [position],
                                           'Press': [(l+1)],
                                           'ADC @ Hardstop 1': [adc_hard_stop_2['Grip ADC']],
                                           'ADC @ Hardstop 2': [adc_hard_stop_4['Grip ADC']],
                                           'Saturated': [Saturated],
                                           '0N ADC': [max_ADC],
                                           '7N ADC': [min_ADC],
                                           'Max Displacement': [Travel]
                                           }

                            df32 = pd.DataFrame(Data_Values)

                            # df33 = df33.append(df32)

                            df33 = pd.concat([df33, df32])    #append going away

                            l = l + 1

                            print('debug')
        df33.to_excel(full_path+'/'+'GTK ' + folders +' results.xlsx', index=False)
    print('Done!')