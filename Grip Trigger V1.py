from tkinter import filedialog
import pandas as pd
import os
import tkinter.filedialog
import openpyxl
from os import walk
import json
import math
import matplotlib.pyplot as plt


Test_Number = 1
mainloop = 0

df33 = pd.DataFrame()

# File dialog to select folder containing excel files to import
full_path = tkinter.filedialog.askdirectory(initialdir='C://')
os.chdir(full_path)

f1 = []
f2 = []
for (dirpath, dirnames, filenames) in walk(full_path):
    f1.extend(filenames)
    f2.extend(dirnames)
    mainloop = mainloop + 1
    if mainloop == 2:
        break


    for folders in f2:
        if folders != 'Plots' or folders != 'Results Package':
            f3 = []
            for (dirpath2, dirnames2, filenames2) in walk(full_path+'/'+folders):
                f3.extend(dirnames2)
            # df1 = pd.DataFrame()
            # df2 = pd.DataFrame()
            print(folders)
            csv = 0
            excel = 0
            report = 0
            MD = 0
            UP = 0
            for folders2 in f3:
                if folders2 != 'Plots' or folders2 != 'Results Package':
                    if csv == 1 and excel == 1 and report == 1 and MD == 1 and UP == 1:
                        break

                    if folders2 == 'position1' or 'position2':
                        if folders2 == 'position1':
                            MD = 1
                        if folders2 == 'position2':
                            UP = 1

                        f4 = []
                        for (dirpath3, dirnames3, filenames3) in walk(full_path+'/'+folders+'/'+folders2):
                            f4.extend(filenames3)
                        print('debug')
                        df1 = pd.DataFrame()
                        df2 = pd.DataFrame()
                        print(folders2)
                        for files in f4:
                            if files.endswith(".csv"):
                                if files.startswith('230'):
                                    test_path = full_path+'/'+folders+'/'+folders2+'/'+files
                                    print(test_path)
                                    df1 = pd.read_csv(test_path)
                                    csv = 1
                            if files.endswith('data.xlsx'):
                                test_path1 = full_path+'/'+folders+'/'+folders2+'/'+files
                                print(test_path1)
                                df2 = pd.read_excel(test_path1, sheet_name='force-displacement CSV data')
                                excel = 1
                            if files.endswith('.json'):
                                json_file = files

                        seconds_list = []
                        print(folders+' '+files)
                        for tick in df1['rtc_ticks']:
                            second_ticks = tick / 125000
                            seconds_list.append(second_ticks)

                        rtc_delta = []
                        rtc_ET = []
                        previous_value = seconds_list[0]
                        for second in seconds_list:

                            second_delta = second - previous_value
                            rtc_delta.append(second_delta)
                            ET = second_delta + previous_value - seconds_list[0]
                            rtc_ET.append(ET)
                            previous_value = second


                        df1.insert(22, 'rtc[s]', seconds_list)
                        df1.insert(23, 'rtc delta', rtc_delta)
                        df1.insert(24, 'rtc ET[s]', rtc_ET)

                        adc_2 = df1['grip_adc']
                        force_1 = df2[df2.columns[0]]
                        adc_5 = []
                        for adc_1 in adc_2:
                            adc_5.append(adc_1)
                        i = 0
                        for adc in adc_5:
                            if adc_5[i] < adc_5[i+1] and adc_5[i+1] < adc_5[i+2] and adc_5[i+2] < adc_5[i+3] and adc_5[i+3] < adc_5[i+4] and adc_5[i+4] < adc_5[i+5] and adc_5[i+5] < adc_5[i+6] and adc_5[i+6] < adc_5[i+7]:
                                break
                            i = i + 1

                        adc_5_list = []
                        adc_5_count = 0
                        prev_min = 14000
                        for adc in adc_5:
                            if adc_5_count > 0 and adc > 12000 and min(adc_5_list) < 10000:
                                break
                            if adc_5_count > 0 and min(adc_5_list) != prev_min:
                                min_index = adc_5_count
                                prev_min = min(adc_5_list)
                            adc_5_count = adc_5_count + 1
                            adc_5_list.append(adc)

                        force_count = 0
                        force_1_list = []
                        prev_max = 1
                        for force in force_1:
                            if force_count > 0 and force < 3 and max(force_1_list) > 6:
                                break
                            if force_count > 0 and max(force_1_list) != prev_max:
                                max_index = force_count
                                prev_max = max(force_1_list)
                            force_count = force_count + 1
                            force_1_list.append(force)


                        force_max_row = df2.iloc[max_index]
                        ADC_max_row = df1.iloc[min_index]

                        time_difference = ADC_max_row['rtc ET[s]'] - force_max_row['time']

                        # index_delta = abs(min_index-max_index)

                        new_time_list = []
                        for time_3 in df2['time']:
                            new_time = time_3 + time_difference
                            new_time_list.append(new_time)




                        reset_force = []
                        for force_value in force_1:
                            reset_force.append(force_value)

                        k = 0
                        for force in force_1:
                            # if force_1[k] < force_1[k+1] and force_1[k+1] < force_1[k+2] and force_1[k+2] < force_1[k+3] and force_1[k+3] < force_1[k+4] and force_1[k+4] < force_1[k+5] and force_1[k+5] < force_1[k+6] and force_1[k+6] < force_1[k+7]:
                            #     break
                            if force > 0.15:
                                break
                            k = k + 1

                        if k == 0 or 1:
                            k = 2

                        force_2 = reset_force[k-2:]

                        trimmed_ADC_1 = df1.loc[df1['rtc ET[s]'] > time_difference]

                        #creates plot to select points for initial travel

                        min_press = max(trimmed_ADC_1['grip_adc'])
                        adclist = []
                        for adc in trimmed_ADC_1['grip_adc']:
                            adc_1000 = (adc/1000)-min_press/1000
                            adclist.append(-1*adc_1000)

                        adclist3 = []
                        for adc2 in trimmed_ADC_1['grip_adc']:
                            adclist3.append(adc2)

                        adc_5_len = len(adc_5)
                        adclist2 = adclist[i:]
                        adclist4 = adclist3[i:]


                        x1 = trimmed_ADC_1['rtc ET[s]']
                        x1_len = len(x1)
                        x11 = x1[i:]

                        # reset_time_list = []
                        # for time in x11:
                        #     reset_time = time - x1[i]
                        #     reset_time_list.append(reset_time)

                        x2 = df2[df2.columns[1]]
                        x2_len = len(x2)




                        x22 = x2[k-2:]

                        reset_time_list2 = []
                        for time2 in x22:
                            reset_time2 = time2 - x22[k-2]
                            reset_time_list2.append(reset_time2)


                        files_2 = os.listdir(full_path+'/'+folders)


                        f = open(full_path+'/'+folders+'/'+folders2+'/'+ json_file)
                        file_info = json.load(f)

                        Unit_SN = file_info['serial_number']




                        if not os.path.exists(full_path+'/Results Package'):
                            os.makedirs(full_path+'/Results Package')
                        if not os.path.exists(full_path+'/Results Package/Plots'):
                            os.makedirs(full_path+'/Results Package/Plots')
                        if not os.path.exists(full_path+'/Results Package/Plots'+Unit_SN+folders2+'.png'):

                            fig1 = plt.figure()
                            plt.plot(new_time_list,force_1, label='Load Cell')
                            plt.plot(x1,adclist, label='Normalised ADC')

                            plt.legend(title='KPI Sources')
                            # plt.show()
                            print('debug')








                            plt.savefig(full_path+'/Results Package/Plots/'+Unit_SN+' '+folders2+'.png',dpi=500)
                            plt.close()


                        df33 = pd.DataFrame()
                        force_4 = df2.iloc[:,0]

                        force_4_len = len(force_4)
                        force_5 = force_4[k:]

                        Fox_Force_5 = []
                        for jj in force_5:
                            Fox_Force_5.append(jj)

                        force_5.index -= (k)

                        # df33['ADC Clock'] = pd.Series(reset_time_list)
                        df33['Fore ADC'] = pd.Series(adclist4)
                        df33['Force Clock'] = pd.Series(reset_time_list2)
                        df33['Force'] = pd.Series(Fox_Force_5)


                        # df33.to_excel(full_path+'/Results Package/'+Unit_SN+' '+folders2+' results.xlsx', index=False)
                        report = 1

print('Done!')








