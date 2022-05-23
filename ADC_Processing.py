from tkinter import filedialog
import os
import tkinter.filedialog
from Fore_Trigger_Post_Processing_V3 import fore_trigger_post_processing_V3
from Grip_Trigger_Post_Processing_V4 import grip_trigger_post_processing_v4
from Pinch_Data_Post_Processing_V5 import pinch_data_post_processing_v5
from Stylus_ADC_V2 import stylus_adc_v2


Test_Number = 1



# File dialog to select folder containing excel files to import
full_path = tkinter.filedialog.askdirectory(initialdir='.')
os.chdir(full_path)


full_pathlower= full_path.lower()

if full_pathlower.__contains__('grip'):
    print('grip')
    # import Grip_Trigger_Sync_V2
    grip_trigger_post_processing_v4(full_path)

if full_pathlower.__contains__('index'):
    fore_trigger_post_processing_V3(full_path)

if full_pathlower.__contains__('pinch'):
    pinch_data_post_processing_v5(full_path)

if full_pathlower.__contains__('stylus'):
    stylus_adc_v2(full_path)
