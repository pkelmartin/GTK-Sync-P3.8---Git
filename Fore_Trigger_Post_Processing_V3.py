import pandas as pd
import os
from os import walk
import matplotlib.pyplot as plt

def fore_trigger_post_processing_V3(full_path):

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
                        df3.insert(1, 'Fore ADC', Fore_ADC)

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
                            max_ADC = max(adc_frame['Fore ADC'])

                            min_ADC = min(adc_frame['Fore ADC'])

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
                                if force_2 > last_value:
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

                                adctime = adc_frame['ADC Clock']

                                adc_list = adc_frame['Fore ADC']

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
                                ax2 = ax1.twinx()
                                ax2.plot(adctime,adc_normal, label='ADC', color='red')
                                ax2.axhline(y=700, color='green', label='700 ADC')
                                # plt.plot(force_clock,z_pos_list)



                                plt.legend(title='Legend')





                                plt.savefig(full_path+'/'+folders+' Plots/'+sn+'-'+position+' Press '+str(l+1)+'.png',dpi=500)
                                plt.close()

                            Data_Values = {'S/N': [sn],
                                           'Side': [side],
                                           'Position': [position],
                                           'Press': [(l+1)],
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