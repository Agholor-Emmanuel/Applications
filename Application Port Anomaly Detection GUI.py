# -*- coding: utf-8 -*-
"""
Created on Mon Dec 27 07:44:45 2021

@author: Emmanuel
"""

import tkinter as tk
import pandas as pd
import numpy as np
from tkinter import ttk
from tkinter.filedialog import askopenfile
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

root = tk.Tk()
root.configure(bg="#C0c0c0")
root.title('Port Scanner Validation')
canvas = tk.Canvas(root, width=1000, height=600, bg="#C0c0c0")
canvas.grid(columnspan=5, rowspan=6)


#Instructions


label_base = tk.Label(root, text="NO FILE SELECTED", font = "none 6 bold", bg="#C0c0c0", height=1, width=50)
label_base.grid(column=0, row=1)


label_new = tk.Label(root, text="NO FILE SELECTED", font = "none 6 bold", bg="#C0c0c0", height=1, width=50)
label_new.grid(column=0, row=3)

frame1 = tk.LabelFrame(root, text="Data", height=580, width=730, bg="#C0c0c0")
frame1.grid(columnspan=3, rowspan=6, column=1, row=0)


#Browser Functions

def open_base():
    browse_base.config(text="Loading...")
    base_file = askopenfilename(initialdir="/", title="Select A File", filetype = [("Excel file", "*.xlsx")])
    if base_file:
        label_base["text"] = base_file
        browse_base.config(text="SELECT BASE SCAN")
            
            
   
def open_new():
      browse_new.config(text="Loading...")
      new_file = askopenfilename(initialdir="/", title="Select A File", filetype = [("Excel file", "*.xlsx")])
      if new_file:
          label_new["text"] = new_file
          browse_new.config(text="SELECT NEW SCAN")
          
         
   
def run():
    compare.config(text="RUNNING")
    base_filepath = label_base["text"]
    new_filepath = label_new["text"]
    df2 = pd.read_excel(base_filepath)
    df3 = pd.read_excel(new_filepath)     
         
    df_check = df2.iloc[:,[0,1,3,4]] 
    df_test = df3.iloc[:,[0,1,3,4]]

    df_result = pd.merge(df_check, df_test, on=['IP', 'Port'], how = 'outer')
         
    df_result['Port changes'] = np.where((df_result['State_x'] == df_result['State_y']), 'No changes', 'Port is ' + df_result['State_y'])
    df_result['Port changes'] = np.where(((df_result['State_x'].isnull()) & (df_result['State_y'].notnull())), 'New Port was ' + df_result['State_y'], df_result['Port changes'])
    df_result['Port changes'] = np.where(((df_result['State_x'].notnull()) & (df_result['State_y'].isnull())), 'Port previously used but now closed ', df_result['Port changes'])
    
    df_result['Service changes'] = np.where((df_result['Service_x'] == df_result['Service_y']), 'No changes', "New Service (" + df_result['Service_y'] + ") on this port")
    df_result['Service changes'] = np.where(((df_result['Service_x'].isnull()) & (df_result['Service_y'].notnull())), "New Service (" + df_result['Service_y'] + ") on this port", df_result['Service changes'])
    df_result['Service changes'] = np.where(((df_result['Service_x'].notnull()) & (df_result['Service_y'].isnull())), 'No Service Running', df_result['Service changes'])

    df_result['State_x'] = df_result['State_x'].fillna('N/A')
    df_result['State_y'] = df_result['State_y'].fillna('N/A')
    df_result['Service_x'] = df_result['Service_x'].fillna('N/A')
    df_result['Service_y'] = df_result['Service_y'].fillna('N/A')
    
    global df
    df = df_result
    
    compare_treeview = ttk.Treeview(frame1)
    compare_treeview.place(relheight=1, relwidth=1)
    
    scrolly = tk.Scrollbar(compare_treeview, orient="vertical", command=compare_treeview.yview)
    scrollx = tk.Scrollbar(compare_treeview, orient="horizontal", command=compare_treeview.xview)
    compare_treeview.configure(xscrollcommand=scrollx.set, yscrollcommand=scrolly.set)    
    scrollx.pack(side="bottom", fill="x")
    scrolly.pack(side="right", fill="y")
    
    compare_treeview["column"] = list(df_result.columns)
    compare_treeview["show"] = "headings"
    for column in compare_treeview["column"]:
        compare_treeview.heading(column, text=column)
        
    df_result_rows =  df_result.to_numpy().tolist()
    for row in df_result_rows:
        compare_treeview.insert("", "end", values=row)
        
    compare.config(text="RUN")    
    
    

def save():
    savefile = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),("All files", "*.*") )) 
    df.to_excel(savefile + ".xlsx", index=False, sheet_name="Report")
    tk.messagebox.showerror("STATUS", "FILE SAVED")                                                 


#Browser Buttons

browse_base = tk.StringVar()
browse_base = tk.Button(root, text="SELECT BASE SCAN", command= lambda:open_base(), font = "none 15 bold", bg="gray50", fg="black", height=2, width=20)
browse_base.config(text="SELECT BASE SCAN")
browse_base.grid(column=0, row=0)

browse_new = tk.StringVar()
browse_new = tk.Button(root, text="SELECT NEW SCAN", command= lambda:open_new(), font = "none 15 bold", bg="gray50", fg="black", height=2, width=20)
browse_new.config(text="SELECT NEW SCAN")
browse_new.grid(column=0, row=2)



#Operations Buttons

compare = tk.StringVar()
compare = tk.Button(root, text="RUN", command= lambda:run(), font = "none 15 bold", bg="gray40", fg="black", height=2, width=20)
compare.config(text="RUN")
compare.grid(column = 0, row = 4)

export = tk.StringVar()
export = tk.Button(root, text="SAVE", command= lambda:save(), font = "none 15 bold", bg="gray40", fg="black", height=2, width=20)
export.config(text="SAVE")
export.grid(column = 0, row = 5)


root.mainloop()

