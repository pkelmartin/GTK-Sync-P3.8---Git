import csv

import pandas as pd
import pyodbc
from tkinter import filedialog
import pandas as pd
import os
import tkinter.filedialog
from os import walk
import json
import math
import matplotlib.pyplot as plt

Test_Number = 1

df1 = pd.DataFrame()

# File dialog to select folder containing excel files to import
full_path = tkinter.filedialog.askdirectory(initialdir='.')
os.chdir(full_path)


f1 = []
f2 = []
for (dirpath, dirnames, filenames) in walk(full_path):
    f1.extend(filenames)
    f2.extend(dirnames)

for folder in f2:

    f3 = []
    for (dirpath2, dirnames2, filenames2) in walk(full_path+'/'+folder):
        f3.extend(filenames2)
        print('debug')

    for file in f3:
        if file.endswith('.mdb'):
            test_path = full_path + '/' +folder + '/' + file
            print(test_path)
            if file.startswith('M'):
                Button = "Menu"
                Side = 'L'
            elif file.startswith('X'):
                Button = 'X'
                Side = 'L'
            elif file.startswith('Y'):
                Button = 'Y'
                Side = 'L'
            elif file.startswith('A'):
                Button = 'A'
                Side = 'R'
            elif file.startswith('B'):
                Button = 'B'
                Side = 'R'
            elif file.startswith('H'):
                Button = 'Home'
                Side = 'R'










            MDB = test_path
            DRV = '{Microsoft Access Driver (*.mdb)}'
            # PWD = 'pw'

            driver = '{Microsoft Access Drive (*.mdb, *.accdb)}'
            filepath = 'C:\\Users\\pkelmartin\\Downloads\\ABXY&HM button test data for POR-1 Main units-20220507\\R\\230YV31D3N000Y\\1.mdb'

            myDataSources = pyodbc.dataSources()
            access_driver = myDataSources['MS Access Database']

            cnxn = pyodbc.connect(driver=access_driver, dbq=filepath, autocommit=True)
            crsr = cnxn.cursor()

            tables_list = list(crsr.tables())
            # for table in tables_list:
            #     print(table)
            #     if table[-3] == 'TestData':
            #         test_data = table



            cur = cnxn.cursor()
            qry = "SELECT * FROM Report"
            df = pd.read_sql(qry, cnxn)
            print('debug')



            col_list = list(df)
            average_list = []
            p_list = []
            for col in col_list:
                if col.startswith('p'):
                    average = sum(df[col])/3
                    average_list.append(average)
                    p_list.append(col)


                    print('debug')
            reduced_list = [col_list[1], col_list[2],col_list,[3],col_list[4],col_list[5],col_list[6], col_list[7]]
            reduced_list_1 = []
            reduced_list_1.append(col_list[0])
            for value in col_list[8:-1]:
                reduced_list_1.append(value)
            #df.drop(columns=reduced_list)

            data_values = {reduced_list_1[0]: 'Average' ,
                           reduced_list_1[1]: [average_list[0]],
                           reduced_list_1[2]: [average_list[1]],
                           reduced_list_1[3]: [average_list[2]],
                           reduced_list_1[4]: [average_list[3]],
                           reduced_list_1[5]: [average_list[4]],
                           reduced_list_1[6]: [average_list[5]],
                           reduced_list_1[7]: [average_list[6]],
                           reduced_list_1[8]: [average_list[7]],
                           reduced_list_1[9]: [average_list[8]],
                           reduced_list_1[10]: [average_list[9]],
                           reduced_list_1[11]: [average_list[10]],
                           reduced_list_1[12]: [average_list[11]],
                           reduced_list_1[13]: [average_list[12]],
                           reduced_list_1[14]: [average_list[13]],
                           reduced_list_1[15]: [average_list[14]],
                           reduced_list_1[16]: [average_list[15]],
                           reduced_list_1[17]: [average_list[16]]
                           }

            df3 = pd.DataFrame(data_values)

            df4 = df[reduced_list_1]

            df5 = pd.concat([df3, df4])

            button_series = []
            SN = []
            side_list = []
            for rows in df5['p1']:
                button_series.append(Button)
                SN.append(folder)
                side_list.append(Side)

            df5.insert(0, 'Button', button_series)
            df5.insert(0, 'Side', side_list)
            df5.insert(0, 'Serial Number', SN)


            df1 = pd.concat([df1, df5])

            print('debug')




df1.to_excel(full_path+'/results.xlsx', index=False)

