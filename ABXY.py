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

MDB = 'C:\\Users\\pkelmartin\\Downloads\\ABXY&HM button test data for POR-1 Main units-20220507\\R\\230YV31D3N000Y\\1.mdb'
DRV = '{Microsoft Access Driver (*.mdb)}'
# PWD = 'pw'

driver = '{Microsoft Access Drive (*.mdb, *.accdb)}'
filepath = 'C:\\Users\\pkelmartin\\Downloads\\ABXY&HM button test data for POR-1 Main units-20220507\\R\\230YV31D3N000Y\\1.mdb'

myDataSources = pyodbc.dataSources()
access_driver = myDataSources['MS Access Database']

cnxn = pyodbc.connect(driver=access_driver, dbq=filepath, autocommit=True)
crsr = cnxn.cursor()

tables_list = list(crsr.tables())
for table in tables_list:
    print(table)
    if table[-3] == 'TestData':
        test_data = table



cur = cnxn.cursor()
qry = "SELECT * FROM Report"
df = pd.read_sql(qry, cnxn)
print('debug')

df.to_excel('C:\\Users\\pkelmartin\\Downloads\\ABXY&HM button test data for POR-1 Main units-20220507\\R\\230YV31D3N000Y\\output.xlsx', index=False)