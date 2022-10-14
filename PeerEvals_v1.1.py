# Import Module
from ast import While
from cProfile import label
from fileinput import close
from multiprocessing.resource_sharer import stop
from pathlib import Path
from tkinter import *
from ttkthemes import themed_tk as tk
import tkinter.font as tkFont
from tkinter import filedialog
from asyncio.windows_events import NULL
from cmath import nan
from ftplib import all_errors
from re import M
from tokenize import group
from openpyxl import workbook
import pandas as pd
import csv
import os
import numpy
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import Levenshtein
pd.options.mode.chained_assignment = None  # default='warn'


#Button commands
def Go():
  
  #Clear Output listbox
  GUIOutput.delete(0,END)

  Inputdb_path = Path_db.get(0)
  GoogleForms_path = Path_GF.get(0)
  Output_path = Path_Output.get(0)

  #Open googleforms.csv and save to a variable
  GoogleForms = pd.read_csv(GoogleForms_path,header=0)
  #Open GoogleForms.csv and only read usernames
  Username = pd.read_csv(GoogleForms_path,header=0,usecols=[1])
  UsernameList = (Username.values).tolist()

  #Open Student Database
  db = pd.read_csv(Inputdb_path,header=0)

  #Create an output database
  output_db = db[['name','login_id']]

  #Save vars
  row = 0
  top_group_len = 0
  UsernameList_track = []
  DuplicateUsername_count = 0
  DuplicateUsername_list = []
  number_of_questions = 6
  memberinputerror_count = 0
  memberinputerror_list = []
  #find the total rows of the spreadsheet
  max_rows = Username.shape[0]
  max_rows_db = db.shape[0]
  while row != max_rows:
    #save the username by each row
    UsernameInput = (UsernameList[row])[0]
    ProcessingUserOutput = str("Processing Username: "+UsernameInput)
    print(ProcessingUserOutput)
    GUIOutput.insert(END,"")
    GUIOutput.insert(END,ProcessingUserOutput)


    try:
      #return row that matches username
      matched_user_group = db.loc[db['login_id']== (UsernameInput),'group_name']
      group_name = ((matched_user_group.values).tolist())[0]
      matched_user_group = db.loc[db['login_id']== (UsernameInput),'name']
      user_fullname = ((matched_user_group.values).tolist())[0]
    except:
      #if name does not exist break out of loop
      InputError = str("Error..."+UsernameInput+" does not exist in the database...")
      print(InputError)
      GUIOutput.insert(END,InputError)
      break
    
    #add Username input to list to track for duplicates
    
    if UsernameInput in UsernameList_track:
        DuplicateUsername = str("Error..."+UsernameInput+" has a duplicate input. Ignoring data...")
        print(DuplicateUsername)
        GUIOutput.insert(END,DuplicateUsername)
        DuplicateUsername_list.append(UsernameInput)
        DuplicateUsername_count += 1
        row += 1
        continue
    UsernameList_track.append(UsernameInput)




    #return all the people with the same group name
    matched_group = db.loc[db['group_name']== group_name]
    matched_group_users = matched_group['name'].values
    group_name_output = str("Group Name:"+str(group_name))
    print(group_name_output)
    GUIOutput.insert(END,group_name_output)
    #Create new columns for output db and update with more columns if a group has more members than previous top
    if len(matched_group_users) > top_group_len:
      for columns in range(len(matched_group_users)-top_group_len):
        if columns == 0 and row == 0:
          output_db.insert(2+columns,"Self Evaluation Score",numpy.nan)
        else:
          db_col_count = len(output_db.columns)
          placeholder_list = list(numpy.repeat([nan], max_rows_db))
          output_db.loc[:,'Group Member #'+str(db_col_count-1)+" Score"] = placeholder_list

    #track the current highest number of group members
    top_group_len = len(matched_group_users)

    #Assign self score to output db
    matched_username = GoogleForms.loc[GoogleForms.iloc[:,1]== (UsernameInput)]
    self_scores_list = matched_username.iloc[:,2:8]
    self_scores_total = self_scores_list.values.sum()
    index_num = db[db['login_id'] == (UsernameInput)].index[0]
    output_db.at[index_num,"Self Evaluation Score"] = float(self_scores_total)
    
    

    for member_no in range(len(matched_group_users)):
      if member_no !=0:
        #Assign scores to group members
        member_score = matched_username.iloc[:,((member_no-1) * (number_of_questions+1) + 9):((member_no-1) * (number_of_questions+1) + (9+number_of_questions))]
        total_member_score = member_score.values.sum()

        #make this to a list and remove the user from the member list
        member_names_db_list = matched_group_users.tolist()
        member_names_db_list.remove(user_fullname)
        print(member_names_db_list)

        #grab the input for username by each column
        member_names_inputs = matched_username.iloc[:,((member_no-1)*7+8)]
        member_names_input = member_names_inputs.tolist()[0]
        
        #Detect if number of member input < group members. If true, print error and break out of loop
        if member_names_inputs.isnull().values.any():
            memberinputerror = str(UsernameInput+ " did not evaluate one or more of their members.")
            print(memberinputerror)
            GUIOutput.insert(END,memberinputerror)
            memberinputerror_list.append(UsernameInput)
            memberinputerror_count += 1
            break

        #insert into output_db if name matches
        matching_member = (process.extract(str(member_names_input),member_names_db_list,limit=1))[0][0]
        matched_member = output_db.loc[output_db['name']==matching_member]


        #find the index number of member in output_db
        member_index = matched_member.index.values[0]

        column_index = 3
        while True:
          #assign member score if the column value is empty, else check next column
          if str(matched_member.iloc[0,column_index]) == "nan":
            #Output member scores
            output_db.iat[member_index,column_index] = total_member_score
            break
          else:
            column_index += 1
            continue
        
        ProcessedMemberOutput = str("Processed Member..."+" Input_name:"+member_names_input+" | Full_name:"+matching_member+" | Score:"+str(total_member_score))       
        print(ProcessedMemberOutput)
        GUIOutput.insert(END,ProcessedMemberOutput)
      else:
        continue

    #print current username in process      
    ProcessedUser = str("Processed Username: "+UsernameInput)
    print(ProcessedUser)
    GUIOutput.insert(END,ProcessedUser)
    row += 1
    continue

  #Add all scores and add new column with totals
  score_averages= output_db.mean(axis=1,numeric_only=True)
  output_db["Average Score"] = score_averages

  #Copy comments column from google forms into output db
  comments_col = GoogleForms.iloc[:,-1]
  output_db = output_db.join(comments_col)   

  # if there are duplicate username inputs raise error at end
  if DuplicateUsername_count > 0:
    DuplicateUsername_count_str = ("Warning...There are "+str(DuplicateUsername_count)+" duplicate input(s):")
    print(DuplicateUsername_count_str)
    print(DuplicateUsername_list)
    GUIOutput.insert(END,"")
    GUIOutput.insert(END,DuplicateUsername_count_str)
    GUIOutput.insert(END,DuplicateUsername_list)

  #if a user did not evaluate all members, raise error at end  
  if memberinputerror_count > 0:
    memberinputerror_count_str = ("Warning...There are "+str(memberinputerror_count)+" user(s) that did not evaluate all their members:")
    print(memberinputerror_count_str)
    print(memberinputerror_list)
    GUIOutput.insert(END,"")
    GUIOutput.insert(END,memberinputerror_count_str)
    GUIOutput.insert(END,memberinputerror_list)

  print("Processing Complete...")
  GUIOutput.insert(END,"")
  GUIOutput.insert(END,"Processing Complete...")
  #Output to xlsx in path
  try:
    output_db.to_excel(Output_path+'\\Peer_Eval_output.xlsx')
  except:
    print("Error... Unable to output to excel file. Please close Peer_Eval_output.xlsx")
    GUIOutput.delete(0,END)
    GUIOutput.insert(END,"Error... Unable to output to excel file. Please close Peer_Eval_output.xlsx")

 
def close_it():
    Path_GF.insert(1,"Quitting Program...")
    Path_GF.delete(0,END)
    quit()


def browseFiles_db():
    file_path = filedialog.askopenfilename()
    if file_path[-4:] != ".csv":
      GUIOutput.delete(0,END)
      GUIOutput.insert(1,"Error...File type invalid. Please select a .csv file.")
    else:
      Path_db.delete(0,END)
      Path_db.insert(1,file_path)


def browseFiles_GF():
    file_path = filedialog.askopenfilename()
    if file_path[-4:] != ".csv":
      GUIOutput.delete(0,END)
      GUIOutput.insert(1,"Error...File type invalid. Please select a .csv file.")
    else:
      Path_GF.delete(0,END)
      Path_GF.insert(1,file_path)

def browseFiles_Output():
    file_path = filedialog.askdirectory()
    Path_Output.delete(0,END)
    Path_Output.insert(1,file_path)

# create root window
root = Tk()
#root = tk.ThemedTk()
# root window title and dimension
root.title("Peer Evaluation Score Calculator")
# Set geometry (widthxheight)
root.geometry('650x450')
root.resizable(False,False)
 
# all widgets will be here
Label1 = Label(root,text="Project Groups.csv path:")
ft = tkFont.Font(family='MS Sans',size=10)
Label1["font"] = ft
Label1.place(x=30,y=10)

Path_db = Listbox(root)
Path_db["borderwidth"] = "1px"
ft = tkFont.Font(family='MS Sans',size=10)
Path_db["font"] = ft
Path_db["fg"] = "#333333"
Path_db["justify"] = "left"
Path_db.place(x=30,y=30,width=510,height=20)


Label2 = Label(root,text="Student Responses.csv path:")
ft = tkFont.Font(family='MS Sans',size=10)
Label2["font"] = ft
Label2.place(x=30,y=60)

Path_GF = Listbox(root)
Path_GF["borderwidth"] = "1px"
ft = tkFont.Font(family='MS Sans',size=10)
Path_GF["font"] = ft
Path_GF["fg"] = "#333333"
Path_GF["justify"] = "left"
Path_GF.place(x=30,y=80,width=510,height=20)

Label3 = Label(root,text="Output folder:")
ft = tkFont.Font(family='MS Sans',size=10)
Label3["font"] = ft
Label3.place(x=30,y=110)

Path_Output = Listbox(root)
Path_Output["borderwidth"] = "1px"
ft = tkFont.Font(family='MS Sans',size=10)
Path_Output["font"] = ft
Path_Output["fg"] = "#333333"
Path_Output["justify"] = "left"
Path_Output.place(x=30,y=130,width=510,height=20)

Browse_db=Button(root)
Browse_db["bg"] = "#f0f0f0"
ft = tkFont.Font(family='MS Sans',size=10)
Browse_db["font"] = ft
Browse_db["fg"] = "#000000"
Browse_db["justify"] = "center"
Browse_db["text"] = "Browse..."
Browse_db.place(x=550,y=30,width=80,height=20)
Browse_db["command"] = browseFiles_db

Browse_GF=Button(root)
Browse_GF["bg"] = "#f0f0f0"
ft = tkFont.Font(family='MS Sans',size=10)
Browse_GF["font"] = ft
Browse_GF["fg"] = "#000000"
Browse_GF["justify"] = "center"
Browse_GF["text"] = "Browse..."
Browse_GF.place(x=550,y=80,width=80,height=20)
Browse_GF["command"] = browseFiles_GF

Browse_Output=Button(root)
Browse_Output["bg"] = "#f0f0f0"
ft = tkFont.Font(family='MS Sans',size=10)
Browse_Output["font"] = ft
Browse_Output["fg"] = "#000000"
Browse_Output["justify"] = "center"
Browse_Output["text"] = "Browse..."
Browse_Output.place(x=550,y=130,width=80,height=20)
Browse_Output["command"] = browseFiles_Output

GUIOutput = Listbox(root)
GUIOutput["borderwidth"] = "1px"
ft = tkFont.Font(family='MS Sans',size=10)
GUIOutput["font"] = ft
GUIOutput["fg"] = "#333333"
GUIOutput["justify"] = "left"
GUIOutput.place(x=30,y=170,width=590,height=230)

scrollbar = Scrollbar(root, orient= 'vertical')
scrollbar.pack(side= RIGHT, fill= BOTH)
GUIOutput.config(yscrollcommand= scrollbar.set)
#Configure the scrollbar
scrollbar.config(command= GUIOutput.yview)

Go_Button=Button(root)
Go_Button["bg"] = "#f0f0f0"
ft = tkFont.Font(family='MS Sans',size=10)
Go_Button["font"] = ft
Go_Button["fg"] = "#000000"
Go_Button["justify"] = "center"
Go_Button["text"] = "Go"
Go_Button.place(x=200,y=415,width=70,height=25)
Inputdb_path = Path_db.get(1)
GoogleForms_path = Path_GF.get(1)
Go_Button["command"] = Go

Exit_Button=Button(root)
Exit_Button["bg"] = "#f0f0f0"
ft = tkFont.Font(family='MS Sans',size=10)
Exit_Button["font"] = ft
Exit_Button["fg"] = "#000000"
Exit_Button["text"] = "Exit"
Exit_Button.place(x=370,y=415,width=70,height=25)
Exit_Button["command"] = close_it



# Execute Tkinter
root.mainloop()
