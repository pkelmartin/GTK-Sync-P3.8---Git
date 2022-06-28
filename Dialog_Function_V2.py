from tkinter import *

from tkinter import messagebox
import tkinter
import tkinter.filedialog


top = Tk()
top.geometry("250x200")




def helloCallBack():
    global full_path
    full_path = tkinter.filedialog.askdirectory(initialdir='.')
    # print(full_path)
    # B.destroy()
    top.destroy()
    return



B = Button(top, text = "Select Top Level Folder", command = helloCallBack)
B.place(x = 50,y = 50)
top.mainloop()
print(full_path)

