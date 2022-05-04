from tkinter import filedialog

import numpy as np
import pandas as pd
import os
import tkinter.filedialog
import openpyxl
from os import walk
import json
import math
import matplotlib.pyplot as plt
from datetime import datetime


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

            f4 = []
            for (dirpath3, dirnames3, filenames3) in walk(full_path+'/'+folders):
                f4.extend(filenames3)
            print('debug')
            df1 = pd.DataFrame()
            df2 = pd.DataFrame()

            for files in f4:
                if files.endswith(".csv"):
                    if files.startswith('230'):
                        test_path = full_path+'/'+folders+'/'+files
                        print(test_path)
                        df1 = pd.read_csv(test_path)
                        csv = 1
                if files.endswith('time.csv'):
                    test_path1 = full_path+'/'+folders+'/'+files
                    print(test_path1)
                    df2 = pd.read_csv(test_path1)
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

            df7 = pd.DataFrame()
            previous_value = 0
            oo = 0
            found = 0
            for adc_3 in df1['stylus']:
                if adc_3 != previous_value:
                    if oo - found > 100:
                        if previous_value == 0:
                            df7 = df7.append(df1.iloc[oo-1])
                            found = oo
                        elif adc_3 == 0:
                            df7 = df7.append(df1.iloc[oo-1])
                            found = oo
                previous_value = adc_3
                oo = oo + 1

            index_adc = df7.index.tolist()

            k = 1
            d = {}
            for index in index_adc:
                if (k % 2) == 0:
                  if k == 2:
                      d[index] = pd.DataFrame()
                      d[index] = df1.iloc[:index]
                  else:
                      d[index] = pd.DataFrame()
                      d[index] = df1.iloc[index_adc[k-3]:index]



                k = k + 1

            adc_2 = df1['stylus']

            Press_Number = 1
            for adc_interval in d:

                adc_df = d[adc_interval]

                Unpressed = min(adc_df['stylus_raw'])
                Pressed = max(adc_df['stylus_raw'])

                files_2 = os.listdir(full_path+'/'+folders)


                f = open(full_path+'/'+folders+'/'+ json_file)
                file_info = json.load(f)

                Unit_SN = file_info['serial_number']

                Data_Values = {'S/N': [Unit_SN],
                               'Press': [Press_Number],
                               'Unpressed ADC': [Unpressed],
                               'Pressed ADC': [Pressed],
                               'ADC Delta': [Pressed - Unpressed]
                               }

                df32 = pd.DataFrame(Data_Values)

                # df33 = df33.append(df32)

                df33 = pd.concat([df33, df32])    #append going away

                Press_Number = Press_Number + 1


df33.to_excel(full_path+'/results.xlsx', index=False)


