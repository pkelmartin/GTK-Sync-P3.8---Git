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

force_col = 0
Z_col = 2

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

                    if folders2 == 'position 1' or 'position 2':
                        if folders2 == 'position 1':
                            MD = 1
                        if folders2 == 'position 2':
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
                                # df2 = pd.read_excel(test_path1, sheet_name='Sheet3')
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

                        adc_2 = df1['fore_adc']
                        force_1 = df2[df2.columns[force_col]]
                        Z_Pos = df2[df2.columns[Z_col]]
                        adc_5 = []
                        for adc_1 in adc_2:
                            adc_5.append(adc_1)
                        i = 0
                        adc_list = []
                        found = 0
                        peak_1 = 0
                        valley_1 = 0
                        for adc in adc_5:
                            if i > 0:
                                max_adc = max(adc_list)
                                if min(adc_list) < (max(adc_list)-4000) and peak_1 == 0:
                                    peak_1 = 1
                                if peak_1 == 1 and (max_adc - adc) < 1000 and valley_1 == 0:
                                    valley_1 = 1
                                if peak_1 == 1 and valley_1 == 1 and (adc - min(adc_list)) < 1000:
                                    break
                                if max_adc - min(adc_list) > 1000:
                                    if abs(max(adc_list) - adc) < 50:
                                        found = 1
                                        break
                            i = i + 1
                            adc_list.append(adc)

                        if found == 0:
                            i = 0
                            peak_1 = 0
                            valley_1 = 0
                            adc_list = []
                            # found = 0
                            for adc in adc_5:
                                if i > 0:
                                    max_adc = max(adc_list)
                                    if min(adc_list) < (max(adc_list)-4000) and peak_1 == 0:
                                        peak_1 = 1
                                    if peak_1 == 1 and (max_adc - adc) < 1000 and valley_1 == 0:
                                        valley_1 = 1
                                    if peak_1 == 1 and valley_1 == 1 and (adc - min(adc_list)) < 1000:
                                        break

                                    if max_adc - min(adc_list) > 1000:
                                        if abs(max(adc_list) - adc) < 100:
                                            found = 1
                                            break
                                i = i + 1
                                adc_list.append(adc)

                        if found == 0:
                            i = 0
                            peak_1 = 0
                            valley_1 = 0
                            adc_list = []
                            # found = 0
                            for adc in adc_5:
                                if i > 0:
                                    max_adc = max(adc_list)
                                    if min(adc_list) < (max(adc_list)-4000) and peak_1 == 0:
                                        peak_1 = 1
                                    if peak_1 == 1 and (max_adc - adc) < 1000 and valley_1 == 0:
                                        valley_1 = 1
                                    if peak_1 == 1 and valley_1 == 1 and (adc - min(adc_list)) < 1000:
                                        break
                                    if max_adc - min(adc_list) > 1000:
                                        if abs(max(adc_list) - adc) < 250:
                                            found = 1
                                            break
                                i = i + 1
                                adc_list.append(adc)

                        if found == 0:
                            i = 0
                            peak_1 = 0
                            valley_1 = 0
                            adc_list = []
                            # found = 0
                            for adc in adc_5:
                                if i > 0:
                                    max_adc = max(adc_list)
                                    if min(adc_list) < (max(adc_list)-4000) and peak_1 == 0:
                                        peak_1 = 1
                                    if peak_1 == 1 and (max_adc - adc) < 1000 and valley_1 == 0:
                                        valley_1 = 1
                                    if peak_1 == 1 and valley_1 == 1 and (adc - min(adc_list)) < 1000:
                                        break
                                    if max_adc - min(adc_list) > 1000:
                                        if abs(max(adc_list) - adc) < 600:
                                            found = 1
                                            break
                                i = i + 1
                                adc_list.append(adc)

                        adc_end_row = df1.iloc[i]


                        reset_force = []
                        for force_value in force_1:
                            reset_force.append(force_value)

                        z_reset = []
                        for z_p in Z_Pos:
                            z_reset.append(z_p)

                        k = 0
                        found_2 = 0
                        peakf_1 = 0
                        valleyf_1 = 0

                        max_force_list = []
                        for force in force_1:
                            if k > 0:
                                if max(max_force_list) > 5 and peakf_1 == 0:
                                    peakf_1 = 1
                                if peakf_1 == 1 and force < 0.5 and valleyf_1 == 0:
                                    valleyf_1 = 1
                                if peakf_1 == 1 and valleyf_1 == 1 and force > 2:
                                    break
                                if max(max_force_list) > 6:
                                    if force < 0:
                                        found_2 = 1
                                        break
                            max_force_list.append(force)
                            k = k + 1

                        if found_2 == 0:
                            k = 0
                            peakf_1 = 0
                            valleyf_1 = 0
                            max_force_list = []
                            for force in force_1:
                                if k > 0:
                                    if max(max_force_list) > 5 and peakf_1 == 0:
                                        peak_1 = 1
                                    if peakf_1 == 1 and force < 0.5 and valleyf_1 == 0:
                                        valley_1 = 1
                                    if peakf_1 == 1 and valleyf_1 == 1 and force > 2:
                                        break
                                    if max(max_force_list) > 6:
                                        if force < 0.05:
                                            break
                                max_force_list.append(force)
                                k = k + 1


                        if found_2 == 0:
                            k = 0
                            peakf_1 = 0
                            valleyf_1 = 0
                            max_force_list = []
                            for force in force_1:
                                if k > 0:
                                    if max(max_force_list) > 5 and peakf_1 == 0:
                                        peak_1 = 1
                                    if peakf_1 == 1 and force < 0.5 and valleyf_1 == 0:
                                        valley_1 = 1
                                    if peakf_1 == 1 and valleyf_1 == 1 and force > 2:
                                        break
                                    if max(max_force_list) > 6:
                                        if force < 0.12:
                                            break
                                max_force_list.append(force)
                                k = k + 1

                        # if k == 0 or 1:
                        #     k = 2

                        force_2 = reset_force[k-2:]

                        force_end_row = df2.iloc[k]

                        print('git hub test')

                        time_difference = adc_end_row['rtc ET[s]']-force_end_row['time']

                        new_time_list = []
                        for time_3 in df2['time']:
                            new_time = time_3 + time_difference
                            new_time_list.append(new_time)

                        #creates plot to select points for initial travel

                        # trimmed_ADC_1 = df1.loc[df1['rtc ET[s]'] > time_difference]

                        min_press = max(df1['fore_adc'])
                        adclist = []
                        for adc in df1['fore_adc']:
                            adc_1000 = (adc/1000)-min_press/1000
                            adclist.append(-1*adc_1000)


                        adclist3 = []
                        for adc2 in df1['fore_adc']:
                            adclist3.append(adc2)

                        adc_5_len = len(adc_5)
                        adclist2 = adclist[i:]
                        adclist4 = adclist3[i:]


                        x1 = df1['rtc ET[s]']
                        x1_len = len(x1)
                        x11 = x1[i:]

                        reset_time_list = []
                        for time in x11:
                            reset_time = time - x1[i]
                            reset_time_list.append(reset_time)

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
                        # Build_info = file_info['board_info']
                        # Build = Build_info['device version']


                        if not os.path.exists(full_path+'/Results Package'):
                            os.makedirs(full_path+'/Results Package')
                        # if not os.path.exists(full_path+'/Results Package/Plots'):
                        #     os.makedirs(full_path+'/Results Package/Plots')
                        # if not os.path.exists(full_path+'/Results Package/Plots'+Unit_SN+folders2+'.png'):
                        #
                        #     fig1 = plt.figure()
                        #     plt.plot(new_time_list,reset_force, label='Load Cell')
                        #     plt.plot(x1,adclist, label='Normalised ADC')
                        #
                        #     plt.legend(title='KPI Sources')
                        #     # plt.show()
                        #     print('debug')
                        #
                        #
                        #
                        #
                        #
                        #
                        #
                        #
                        #     plt.savefig(full_path+'/Results Package/Plots/'+Unit_SN+' '+folders2+'.png',dpi=500)
                        #     plt.close()


                        df33 = pd.DataFrame()
                        force_4 = df2.iloc[:,0]

                        force_4_len = len(force_4)
                        force_5 = force_4[k:]

                        Fox_Force_5 = []
                        for jj in force_5:
                            Fox_Force_5.append(jj)

                        force_5.index -= (k)

                        if len(x1) > len(new_time_list):

                            df33['ADC Clock'] = pd.Series(x1)
                            df33['Fore ADC'] = pd.Series(adclist3)
                            df33['Force Clock'] = pd.Series(new_time_list)
                            df33['Force'] = pd.Series(reset_force)
                            df33['Z_Position'] = pd.Series(z_reset)

                        else:
                            df33['Force Clock'] = pd.Series(new_time_list)
                            df33['Force'] = pd.Series(reset_force)
                            df33['Z_Position'] = pd.Series(z_reset)

                            jj = pd.Series(adclist3)

                            df33.insert(0, 'Fore ADC', jj)
                            df33.insert(0, 'ADC Clock', x1)






                        df33.to_excel(full_path+'/Results Package/'+Unit_SN+' '+folders2+' results.xlsx', index=False)
                        report = 1

print('Done!')








