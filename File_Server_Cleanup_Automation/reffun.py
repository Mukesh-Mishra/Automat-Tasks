import datetime
from tkinter.filedialog import askopenfilename
import sqlite3
import pandas as pd
import os
from posixpath import curdir
from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
import os
from turtle import up
import re
import numpy as np
import pandas as pd
import gc
import datetime
import csv
import sqlite3
from tkinter import simpledialog
dbname = 'mydb.db'
# #pyinstaller --onefile --icon=img.ico -w Converter.py
#=============Welcome function==========================================
def welcomemsg():
    now = datetime.datetime.now()
    hour = now.hour
    if 5 <= hour < 12:
        greeting = "Good Morning"
    elif hour < 18:
        greeting = "Good Afternoon"
    else:
        greeting = "Good Evening"
    return greeting

def getfile_fun():
    filename = askopenfilename(initialdir="Desktop/",title='Choose an excel file')
    #filename = askopenfilename(initialdir="Desktop/",filetypes=[('All Files',"*."), ("Excel file (.*xlsx)","*.xlsx"), ('CSV files',"*.csv")],title='Choose an excel file')
    #filename = askopenfilename(parent=root,mode='r',filetypes=[("Excel file","*.xlsx",".*xlx")],title='Choose an excel file')
    return filename

def add_in_tbl(table_name,data_frame):
    conn=sqlite3.connect(dbname)
    try:
        cur = conn.cursor()
        cur.execute('''DROP TABLE IF EXISTS table_name''')
        data_frame.to_sql(table_name, conn, if_exists='replace', index=False)
    finally:
        conn.close()

def read_tbl(table_name):
    conn=sqlite3.connect(dbname)
    try:
        data_frame = pd.read_sql('select * from {}'.format(table_name), conn)
        return data_frame
    finally:
        conn.close()

def list_tbl():
    conn=sqlite3.connect(dbname)
    try:
        sql_query = """SELECT name FROM sqlite_schema WHERE type ='table' AND name NOT LIKE 'sqlite_%'"""
        cursor = conn.cursor()
        cursor.execute(sql_query)
        all_tables = cursor.fetchall()
        return all_tables
    finally:
        conn.close()

def folder_path_dump(pathtxtfile):
    #pathtxtfile = os.path.normpath(os.path.expanduser("~/Desktop")) + '\\shared_folders_path.txt'
    df_folder_path=pd.read_csv(pathtxtfile, sep=" ", header=None,  names=["Path_Name"])
    df_folder_path["Path_Name"]=df_folder_path["Path_Name"].str.strip()
    add_in_tbl('tbl_paths', df_folder_path)

    list_folder_path = df_folder_path["Path_Name"].tolist()
    df_folder_path_dump=pd.DataFrame(columns=["SamAccountName", "Path_Name"])
    dumpdate = (datetime.date.today()).strftime('%d-%m-%Y')
    desktop_path = os.path.normpath(os.path.expanduser("~/Desktop"))
    for path in list_folder_path:
        path1 = path + '/'
        try:
            subfolders = [ f.name for f in os.scandir(path1) if f.is_dir() ]
            pattern = re.compile('^[A,a,B,b]')
            subfolders_list = [ s for s in subfolders if (pattern.match(s) and len(s)==8) ]
            subfolders_dict= {"SamAccountName": subfolders_list,"Path_Name": path}

            df_subfolders=pd.DataFrame.from_dict(subfolders_dict)
            df_folder_path_dump = pd.concat([df_folder_path_dump,df_subfolders], axis=0)
        except Exception as e:
            logf = open(desktop_path+ "\\get_folders_path_logs.log", "a")
            logf.write("Failed: Path {0}: {1}\n".format(str(path), str(e)))  

    df_folder_path_dump["Path_Name"] = df_folder_path_dump.groupby("SamAccountName")["Path_Name"].transform(lambda x: ", ".join(x.unique()))
    df_folder_path_dump = df_folder_path_dump.drop_duplicates(subset=["SamAccountName"], keep='first')
    df_folder_path_dump.reset_index(inplace=False)
    
    df_folder_path_dump.to_csv(desktop_path + '\\' + 'folder-path-dump' + "-" + dumpdate + '.csv', index=False)
    messagebox.showinfo("Successful", "Hey! Folders name saved successfully on Desktop.!")

def reanme_paths_folders(input_file):
    
    #input_File= lblfilename_text
    #input_file = os.path.normpath(os.path.expanduser("~/Desktop")) + '\\Inactive_Users.csv'
    df_input=pd.read_csv(input_file)
    all_folders_list=df_input["SamAccountName"].tolist()
  
    df_path=read_tbl('tbl_paths')
    all_path_list=df_path["Path_Name"].tolist()

    desktop_path_for_log = os.path.normpath(os.path.expanduser("~/Desktop"))
                       
    for path in all_path_list:
        if not os.path.exists(path): #path not exist:
            logf = open(desktop_path_for_log+ "\\Folders_Rename_logs.log", "a")
            logf.write("Failed: {0}: Path does not exit. \n".format(str(path)))
        else:
            try:
                for folder in all_folders_list:
                    try:
                        newname = folder + "_old"
                                                    
                        if not os.path.isdir(os.path.join(path, folder)):
                            logf = open(desktop_path_for_log+ "\\Folders_Rename_logs.log", "a")
                            logf.write("Failed: {0}: Folder does not exit in {1}\n".format(str(folder),str(path)))
                            continue # Not a directory
                            
                        else:  
                            os.rename(os.path.join(path,folder), os.path.join(path,newname)) 
                            logf = open(desktop_path_for_log+ "\\Folders_Rename_logs.log", "a")
                            logf.write("Success: {0}: Folder Renamed....in {1} \n".format(str(folder), str(path)))  

                    except Exception as e:
                            logf = open(desktop_path_for_log+ "\\Folders_Rename_logs.log", "a")
                            logf.write("Failed: {0}: To rename {1} because {2}\n".format(str(folder), str(path),str(e)))
            except Exception as e:
                logf = open(desktop_path_for_log+ "\\Folders_Rename_logs.log", "a")
                logf.write("Failed: {0}: {1}\n".format(str(path), str(e)))
