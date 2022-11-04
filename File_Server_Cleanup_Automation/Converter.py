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
import reffun
import datetime
import csv
import sqlite3
from tkinter import simpledialog

class frmuploader():
    def __init__(self, uploader):
        self.uploader = uploader
        uploader.geometry("900x200+250+100")
        uploader.maxsize(680, 160)
        uploader.minsize(680, 160)
        uploader.iconbitmap()
        uploader['bg'] = 'white'
        uploader.overrideredirect(1)

        self.lblheader = Label(uploader, text=" CONVERT DOCUMENT", borderwidth=2, bg='#4682B4', anchor='w', fg='white',
                          font='Verdana 10 bold')
        self.lblheader.place(x=0, y=0, width=900, height=30)

        self.lbldoctype = Label(uploader, text="Document Type:", font='verdana 8 bold', bg='white')
        self.lbldoctype.place(x=30, y=60, anchor='w')
        self.lblfileselect = Label(uploader, text="Select a File:", font='verdana 8 bold', bg='white')
        self.lblfileselect.place(x=30, y=110, anchor='w')

        # self.lblpatheselect = Label(uploader, text="Select Drive Path:", font='verdana 8 bold', bg='white')
        # self.lblpatheselect.place(x=30, y=140, anchor='w')


        self.combodoctype = ttk.Combobox(uploader, height=15, justify='left', width=40,
                                    values=["  -- Select Document --", "Machines-Dump", "Membership", 
                                    "PoolWise-Lookup", "EntitlementWise-Lookup","Get-Folder-Name","Rename-Folders",
                                    "Folder-Path-Dump","Rename-All-Paths-Folders"],
                                    state="readonly")
        
        self.combodoctype.bind('<<ComboboxSelected>>', self.combobox_item_modified)    
        self.combodoctype.pack()
        self.combodoctype.current(0)
        self.combodoctype.place(x=180, y=50, height=25)

        # -----------------------------------------------------------------------------------------------------------------------------------------------

        self.btnbrowse = Button(uploader, text="Browse", relief="raised", borderwidth=1, bg='gray', fg='white',
                              font='verdana 10 bold',command=self.btnBrowse)
        self.btnbrowse.place(x=180, y=99, width=70, height=26)
        self.btnbrowse.focus_set()

        # self.btnpath = Button(uploader, text="Browse Path", relief="raised", borderwidth=1, bg='gray', fg='white',
        #                       font='verdana 7 bold',command=self.btnBrowse)
        # self.btnpath.place(x=180, y=130, width=75, height=20)
      

        self.lblfilename = Label(uploader, text='No file chosen', font='vardana 8 bold', anchor='w', relief='solid',
                            borderwidth=1)
        self.lblfilename.place(x=250, y=100, width=240, height=25)

        # self.lblpathname = Label(uploader, text='No file chosen', font='vardana 8 bold', anchor='w')
        # self.lblpathname.place(x=250, y=130, width=240, height=25)

        # self.btnupload = Button(uploader, text="Upload", relief="ridge", borderwidth=0, bg='#4682B4', fg='white',
        #                       font='verdana 10 bold',command = lambda : messagebox.askyesno( 'Confirm', 'Do you want to save?'))
        self.btnupload = Button(uploader, text="Convert", relief="ridge", borderwidth=0, bg='#4682B4', fg='white',
                              font='verdana 10 bold',command =self.convert_files )
        self.btnupload.place(x=500, y=100, width=70, height=25)

        self.btnclose = Button(uploader, text="Close", command=self.close_window, relief="solid", borderwidth=0, bg='#CD5C5C',
                             fg='white',
                             font='verdana 10 bold')
        self.btnclose.place(x=575, y=100, width=70, height=25)

   
    def close_window(self):
        self.uploader.destroy()

    def combobox_item_modified (self, event) :
        if self.combodoctype.get() == "Get-Folder-Name":
            self.lblfileselect['text'] = "Drive Location: "
            self.btnupload['text'] = "Get"
        elif self.combodoctype.get() == "Rename-Folders":
            self.lblfileselect['text'] = "Select Folders Name:"
            self.btnupload['text'] = "Rename"
        else:
            self.lblfileselect['text'] = "Select a File:"
            self.btnupload['text'] = "Convert"

    # ---------< file selection>---------------------------------------------------------
    def btnBrowse(self):
        if self.btnupload['text'] in ["Convert","Rename"]:
            filename = reffun.getfile_fun()
            # if file_name != None:
            self.lblfilename.config(text=filename)
        elif self.btnupload['text'] == "Get":
            currdir = os.getcwd()
            foldername = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
            self.lblfilename.config(text=foldername)


    def get_file_path(self,filename):
        # ''' - This gets the full path...file and terminal need  to be in same directory - '''
        file_path = os.path.join(os.getcwd(), filename)
        return file_path

  
    def convert_files(self):
        if self.combodoctype.get() in [""," ","  -- Select Document --"]:
            messagebox.showwarning("Document_Error", "Hey! Select a Document Type.!")
        else:
            try:
                lblfilename_text = self.lblfilename["text"]
                if lblfilename_text == 'No file chosen' or lblfilename_text == '':
                    messagebox.showinfo("", "Hey!  Please Select a File.")
                    self.btnbrowse.focus()

                else: 

                    # if '.xlsx' in lblfilename_text:
                    #     current_Directory = os.path.dirname(os.path.abspath(lblfilename_text))
                    #     os.makedirs(current_Directory, exist_ok=True)
                    # elif '.XLSX' in lblfilename_text:
                    #     current_Directory = os.path.dirname(os.path.abspath(lblfilename_text))
                    #     os.makedirs(current_Directory, exist_ok=True)
                    # elif '.csv' in lblfilename_text:
                    #     current_Directory = os.path.dirname(os.path.abspath(lblfilename_text))
                    #     os.makedirs(current_Directory, exist_ok=True)
                    # else:
                    #     messagebox.showwarning("File_Error", "Hey! You are selecting incorrect file.!")
                    
                    # Membership Dump
                    if self.combodoctype.get() == "Membership":
                        input_file_membership = lblfilename_text

                        df = pd.read_csv(input_file_membership)
                        df["GroupName"] = df.groupby("SamAccountName")["GroupName"].transform(lambda x: ", ".join(x.unique()))
                        df = df.drop_duplicates(subset=["SamAccountName"], keep='first')
                        df.columns = df.columns.str.replace('SamAccountName', 'OLM_ID')
                        df = df.drop(["Name","Title"], axis=1)
                        df.reset_index(inplace=False)
                        dumpdate = (datetime.date.today()).strftime('%d-%m-%Y')
                        df.to_excel(os.path.dirname(os.path.abspath(input_file_membership)) + '\\' + '_Membership-' + dumpdate + '.xlsx', index=False)
                        reffun.add_in_tbl('tbl_membership', df)
                        messagebox.showinfo("Successful", "Hey! Membership File converted successfully.!")
                    elif self.combodoctype.get() == "Machines-Dump":
                        # Machine Dump
                        input_file_machine = lblfilename_text
                        df1 = pd.read_excel(input_file_machine, usecols = ["Assigned User","Desktop Pool","Status"])
                        df1["Desktop Pool"] = df1.groupby("Assigned User")["Desktop Pool"].transform(lambda x: ", ".join(x.unique()))
                        df1 = df1.drop_duplicates(subset=["Assigned User"], keep='first')
                        df1 = df1[df1["Status"].str.contains("Maintenance Mode|Unassigned User Connected") == False]
                        df1["Assigned User"] = df1["Assigned User"].str[-8:]
                        df1 = df1.reindex(columns = ["Assigned User","Desktop Pool","Status"])
                        df1.reset_index(inplace=False)
                        df1.columns = df1.columns.str.replace('Assigned User', 'OLM_ID')
                        dumpdate1 = (datetime.date.today()).strftime('%d-%m-%Y')
                        df1.to_excel(os.path.dirname(os.path.abspath(input_file_machine)) + '\\' + '_Machine_Dump-' + dumpdate1 + '.xlsx', index=False)
                        reffun.add_in_tbl('tbl_machine_dump', df1)
                        messagebox.showinfo("Successful", "Hey! Machine-Dump File converted successfully.!")
         
                    elif self.combodoctype.get() == "PoolWise-Lookup":
                        # Desktop Pool Wise lookup
                        input_file_desktoppoolMachines = lblfilename_text
                        df2 = pd.read_excel(input_file_desktoppoolMachines, usecols = ["Assigned User","Machine","Status"])
                        df2 = df2[df2["Status"].str.contains("Maintenance Mode|Unassigned User Connected") == False]
                        df2["Assigned User"] = df2["Assigned User"].str[-8:]
                        df2 = df2.reindex(columns = ["Assigned User","Machine","Status"])
                        df2.columns = df2.columns.str.replace('Assigned User', 'OLM_ID')
                        df2.reset_index(inplace=False)

                        df1 = reffun.read_tbl('tbl_machine_dump')
                        df = reffun.read_tbl('tbl_membership')
                        df2 = pd.merge(df2, df1, on = "OLM_ID", how = "left")
                        df2 = pd.merge(df2, df, on = "OLM_ID", how = "left")
                        df2 = df2.drop(['Status_y'], axis=1)
                        df2.reset_index(inplace=False)
                        df2.columns = df2.columns.str.replace('Status_x', 'Status')
                        df2 = df2.reindex(columns = ["OLM_ID", "Department", "Machine", "Desktop Pool", "GroupName", "Status"])
                        
                        dumpdate2 = (datetime.date.today()).strftime('%d-%m-%Y')
                        df2.to_excel(os.path.dirname(os.path.abspath(input_file_desktoppoolMachines)) + '\\' + '_Pool-' + dumpdate2 + '.xlsx', index=False)
 
                        messagebox.showinfo("Successful", "Hey! Pool File converted successfully.!")


                    elif self.combodoctype.get() == "EntitlementWise-Lookup":
                        # Desktop Pool Wise lookup
                        input_file_entitlements = lblfilename_text
                        df3 = pd.read_excel(input_file_entitlements, usecols = ["Name","Sessions"])
                        df3["Name"] = df3["Name"].apply(lambda x: x.split("@")[0])
                        df3 = df3.reindex(columns = ["Name","Sessions"])
                        df3.columns = df3.columns.str.replace('Name', 'OLM_ID')
                        df3.reset_index(inplace=False)
                        

                        df1 = reffun.read_tbl('tbl_machine_dump')
                        df = reffun.read_tbl('tbl_membership')
                        df3 = pd.merge(df3, df1, on = "OLM_ID", how = "left")
                        df3 = pd.merge(df3, df, on = "OLM_ID", how = "left")
                        df3 = df3.drop(['Sessions'], axis=1)
                        df3.reset_index(inplace=False)
                        df3 = df3.reindex(columns = ["OLM_ID", "Department","GroupName", "Desktop Pool", "Status"])
                        
                        dumpdate3 = (datetime.date.today()).strftime('%d-%m-%Y')
                        df3.to_excel(os.path.dirname(os.path.abspath(input_file_entitlements)) + '\\' + '_Entitlement-' + dumpdate3 + '.xlsx', index=False)
 
                        messagebox.showinfo("Successful", "Hey! Entitlement File converted successfully.!")

                    elif self.combodoctype.get() == "Get-Folder-Name":
                        input_Folders_Name = lblfilename_text

                        My_directory = input_Folders_Name + '/'
                        subfolders = [ f.name for f in os.scandir(My_directory) if f.is_dir() ]


                        pattern = re.compile('^[A,a,B,b]')
                        final_list = [ s for s in subfolders if (pattern.match(s) and len(s)==8) ]    

                        df4 = pd.DataFrame(final_list, columns=["SamAccountName"])               
 


                        dumpdate4 = (datetime.date.today()).strftime('%d-%m-%Y')
                        desktop_path = os.path.normpath(os.path.expanduser("~/Desktop"))
                        input_Dir_Name = input_Folders_Name.rsplit('/', 1)[-1]
                        df4.to_excel(desktop_path + '\\' + input_Dir_Name + "-" + dumpdate4 + '.xlsx', index=False)
                        messagebox.showinfo("Successful", "Hey! Folders name saved successfully on Desktop.!")


                    elif self.combodoctype.get() == "Rename-Folders":
                        input_File= lblfilename_text
                        print(input_File)
                        df5=pd.read_csv(input_File)
                        all_folders_list=df5["SamAccountName"].tolist()
  
                        basedir = simpledialog.askstring(title="Test",prompt="Enter The Drive Full Path For Renaming The Folders as Foldername_old:")
                        #basedir = rf'{basedir1}'.strip()
                        print(basedir)

                        desktop_path_for_log = os.path.normpath(os.path.expanduser("~/Desktop"))
                        print(desktop_path_for_log)
                        for fn in all_folders_list:
                            try:
                                newname = fn + "_old"
                                                            
                                if not os.path.isdir(os.path.join(basedir, fn)):
                                    logf = open(desktop_path_for_log+ "\\Folder_Rename_logs.log", "a")
                                    logf.write("Failed: {0}: Path does not exit. \n".format(str(fn)))
                                    continue # Not a directory
                                    
                                else:  
                                    os.rename(os.path.join(basedir,fn), os.path.join(basedir,newname)) 
                                    logf = open(desktop_path_for_log+ "\\Folder_Rename_logs.log", "a")
                                    logf.write("Success: {0}: Folder Renamed. \n".format(str(fn)))  

                            except Exception as e:
                                    logf = open(desktop_path_for_log+ "\\Folder_Rename_logs.log", "a")
                                    logf.write("Failed: {0}: {1}\n".format(str(fn), str(e)))

                        messagebox.showinfo("Successful", "Hey! Done")   
                    elif self.combodoctype.get() == "Folder-Path-Dump":
                        input_Folder_Path= lblfilename_text
                        reffun.folder_path_dump(input_Folder_Path)

                    elif self.combodoctype.get() == "Rename-All-Paths-Folders":
                        input_Folder_Path= lblfilename_text
                        reffun.reanme_paths_folders(input_Folder_Path)
                        messagebox.showinfo("Successful", "Hey! Done")


            # except Exception as e:
            #     messagebox.showerror('Error!',e)
            finally:
               # del (df,df1,df2)
                gc.collect()
                

root = tk.Tk()
app = frmuploader(root)
root.mainloop()
