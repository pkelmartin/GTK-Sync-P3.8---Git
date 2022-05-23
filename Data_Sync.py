from tkinter import filedialog
import os
import tkinter.filedialog
from Fore_Trigger_V5 import fore_trigger_V5
from Grip_Trigger_Sync_V2 import grip_trigger_sync_v2
from Pinch_Data_Sync_V4 import pinch_data_sync_v4



print('select folder to be processed')

# df33 = pd.DataFrame()

# File dialog to select folder containing excel files to import
full_path = tkinter.filedialog.askdirectory(initialdir='.')
os.chdir(full_path)

full_pathlower= full_path.lower()

if full_pathlower.__contains__('grip'):
    print('grip')
    # import Grip_Trigger_Sync_V2
    grip_trigger_sync_v2(full_path)
    # Grip_Trigger_Sync_V2
if full_pathlower.__contains__('index'):
    fore_trigger_V5(full_path)
    # Fore_Trigger_V5
if full_pathlower.__contains__('pinch'):
    pinch_data_sync_v4(full_path)
