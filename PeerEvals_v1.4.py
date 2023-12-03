# Import Module
from ast import While
from cProfile import label
from fileinput import close
from multiprocessing.resource_sharer import stop
from pathlib import Path
from tkinter import *
import tkinter.font as tkFont
from tkinter import filedialog
from asyncio.windows_events import NULL
from cmath import nan
from ftplib import all_errors
from re import M
from tokenize import group
import pandas
import csv
import os
import numpy
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import Levenshtein
import openpyxl
import string
from openpyxl import Workbook
from openpyxl.comments import Comment


pandas.options.mode.chained_assignment = None  # default='warn'

 
#Button commands
def Run():
  
  #Clear Output listbox
  GUIOutput.delete(0,END)
  Inputdb_path = Path_db.get(0)
  GoogleForms_path = Path_GF.get(0)
  Output_path = Path_Output.get(0)

  #Open googleforms.csv and save to a variable
  try:
    GoogleForms = pandas.read_csv(GoogleForms_path,header=0, encoding='cp1252')
    Username = pandas.read_csv(GoogleForms_path,header=0,usecols=[1], encoding='cp1252')
    UsernameList = (Username.values).tolist()
  except Exception as e:
    GUIOutput.insert(END,"Student Responses.csv path is required...")
  #Open GoogleForms.csv and only read usernames
  
  

  #Open Student Database (Project Groups)
  try:
    db = pandas.read_csv(Inputdb_path,header=0)
  except:
    GUIOutput.insert(END,"Project Groups.csv path is required...")

  #Create an output database
  output_db = db[['name','login_id']]

  #Save vars
  row = 0
  top_group_len = 0
  UsernameList_track = []
  DuplicateUsername_count = 0
  DuplicateUsername_list = []
  #assign no of questions based on UI input
  number_of_questions = int(clicked1.get())
  memberinputerror_count = 0
  memberinputerror_list = []
  low_match_count = 0
  max_score = 0
  low_match_list = []
  low_match_TF = FALSE
  #find the total rows of the spreadsheet
  max_rows = Username.shape[0]
  max_rows_db = db.shape[0]
  #make a list to store comments
  comment_index_list = []
  comment_list = []
  #Store output
  output_list = []
  uppercase_abc = list(string.ascii_uppercase)
  score_comment_list = []
  while row != max_rows:
    #save the username by each row
    UsernameInput = (UsernameList[row])[0]
    ProcessingUserOutput = str("Processing Username: "+UsernameInput)
    print(ProcessingUserOutput)
    GUIOutput.insert(END,"")
    GUIOutput.insert(END,ProcessingUserOutput)
    output_list.append("")
    output_list.append(ProcessingUserOutput)
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
      output_list.append(InputError)
      break
    
    #add Username input to list to track for duplicates
    
    if UsernameInput in UsernameList_track:
        DuplicateUsername = str("Error..."+UsernameInput+" has a duplicate input. Ignoring data...")
        print(DuplicateUsername)
        GUIOutput.insert(END,DuplicateUsername)
        output_list.append(DuplicateUsername)
        DuplicateUsername_list.append(UsernameInput)
        DuplicateUsername_count += 1
        row += 1
        continue
    UsernameList_track.append(UsernameInput)




    #return all the people with the same group name
    matched_group = db.loc[db['group_name']== group_name]
    matched_group_users = matched_group['name'].values

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
    self_scores_list = matched_username.iloc[:,2: (number_of_questions + 2)]
    self_scores_total = self_scores_list.values.sum()
    index_num = db[db['login_id'] == (UsernameInput)].index[0]
    output_db.at[index_num,"Self Evaluation Score"] = float(self_scores_total)
    
    #remember the max score given to self, used for calculation in computed total score
    if max_score < self_scores_total:
      max_score = self_scores_total

    #remember the number of group members for each user input
    current_group_len = len(matched_group_users)
    #save the names of matched group members, and remove the user from the member list
    member_names_db_list = matched_group_users.tolist()
    member_names_db_list.remove(user_fullname)

    #Output user group and score
    user_output = str("Group Name:"+str(group_name)+" | Full_Name:"+str(user_fullname)+ " | Self_score:"+str(self_scores_total))
    print(user_output)
    GUIOutput.insert(END, user_output)
    output_list.append(user_output)
    
    for member_no in range(current_group_len):
      if member_no !=0:
        #Assign scores to group members
        member_score = matched_username.iloc[:,((member_no-1) * (number_of_questions+1) + (number_of_questions+3)):((member_no-1) * (number_of_questions+1) + ((number_of_questions + 3) +number_of_questions))]
        print(member_score)
        total_member_score = member_score.values.sum()

        
        #grab the input for username by each column
        member_names_inputs = matched_username.iloc[:,((member_no-1)*(number_of_questions+ 1) + (number_of_questions + 2))]
        member_names_input = member_names_inputs.tolist()[0]
        
        #Detect if number of member input < group members. If true, print error and break out of loop
        if member_names_inputs.isnull().values.any():
            memberinputerror = str(UsernameInput+ " did not evaluate one or more of their members.")
            print(memberinputerror)
            GUIOutput.insert(END,memberinputerror)
            output_list.append(memberinputerror)
            memberinputerror_list.append(UsernameInput)
            memberinputerror_count += 1
            break

        #insert into output_db if name matches
        matching_member = (process.extract(str(member_names_input),member_names_db_list,limit=1))[0][0]
        matching_score = (process.extract(str(member_names_input),member_names_db_list,limit=1))[0][1]
        matched_member = output_db.loc[output_db['name']==matching_member]

        #remove from member from member names list after assigning score
        member_names_db_list.remove(matching_member)

        if matching_score <= int(clicked2.get()):
          low_match = str("Warning... low matching score ("+str(matching_score)+") for member: "+ str(matching_member))
          print(low_match)
          GUIOutput.insert(END,low_match)
          output_list.append(low_match)
          low_match_TF = TRUE


        #find the index number of member in output_db
        member_index = matched_member.index.values[0]

        column_index = 3
        while True:
          #assign member score if the column value is empty, else check next column
          if str(matched_member.iloc[0,column_index]) == "nan":
            #Output member scores
            output_db.iat[member_index,column_index] = total_member_score
            cell = (uppercase_abc[column_index+1]+str(member_index+2))
            score_comment_list.append([cell, user_fullname])
            break
          else:
            column_index += 1
            continue
        
        ProcessedMemberOutput = str("Processed Member..."+" Input_name:"+member_names_input+" | Full_name:"+matching_member+" | Score:"+str(total_member_score))       
        print(ProcessedMemberOutput)
        GUIOutput.insert(END,ProcessedMemberOutput)
        output_list.append(ProcessedMemberOutput)
      else:
        continue

    #if there are any low matches (<85 score) append to low_match_list
    if low_match_TF == TRUE:
      low_match_list.append(UsernameInput)
      low_match_count += 1
      #set low match back to FALSE if previously set as TRUE
      low_match_TF = FALSE

    #Store list of comment index and comments
    comment_index_list.append(index_num)

    comment = str(matched_username.iloc[:,-1].values[0])
    comment_list.append(comment)

    #print current username in process      
    ProcessedUser = str("Processed Username: "+UsernameInput)
    print(ProcessedUser)
    GUIOutput.insert(END,ProcessedUser)
    output_list.append(ProcessedUser)
    row += 1
    continue

  #Add all scores and add new column with totals
  score_averages= output_db.mean(axis=1,numeric_only=True)
  output_db["Average Score"] = score_averages


  #Assign comments of each row based on both lists
  output_db["Comments"] = ""

  #Remove comments that do not have values
  for comment_index in range(len(comment_list)):
    if comment_list[comment_index] == 'nan':
      comment_list[comment_index] = ""

  #assign comments based on list of index and comments
  for x in range(len(comment_index_list)):
    output_db.at[comment_index_list[x], "Comments"] = comment_list[x]

  # if there are duplicate username inputs raise error at end
  if DuplicateUsername_count > 0:
    DuplicateUsername_count_str = ("Warning...There are "+str(DuplicateUsername_count)+" user(s) with duplicate inputs:")
    print(DuplicateUsername_count_str)
    print(DuplicateUsername_list)
    GUIOutput.insert(END,"")
    GUIOutput.insert(END,DuplicateUsername_count_str)
    GUIOutput.insert(END,DuplicateUsername_list)
    output_list.append(DuplicateUsername_count_str)
    output_list.append(DuplicateUsername_list)

  #if a user did not evaluate all members, raise error at end  
  if memberinputerror_count > 0:
    memberinputerror_count_str = ("Warning...There are "+str(memberinputerror_count)+" user(s) that did not evaluate all their members:")
    print(memberinputerror_count_str)
    print(memberinputerror_list)
    GUIOutput.insert(END,"")
    GUIOutput.insert(END,memberinputerror_count_str)
    GUIOutput.insert(END,memberinputerror_list)
    output_list.append("")
    output_list.append(memberinputerror_count_str)
    output_list.append(memberinputerror_list)

  #if low match count more than one generate error at the end
  if low_match_count > 1:
    low_match_str = ("Warning...There are "+str(low_match_count)+" user(s) that has matching scores <="+clicked2.get()+"%"+" for members:")
    print(low_match_str)
    print(low_match_list)
    GUIOutput.insert(END,"")
    GUIOutput.insert(END,low_match_str)
    GUIOutput.insert(END,low_match_list)
    output_list.append("")    
    output_list.append(low_match_str)
    output_list.append(low_match_list)

  #Check if total(max) score input is empty or int, if it is int, compute the average score column based on user input
  if total_label_box.get() == "":
    pass
  else:
    try:
      total_score = int(total_label_box.get())
      output_db["Average Score"] = output_db["Average Score"]/max_score * total_score
    except:
      total_score_error = ("Error... Please enter a valid number... Ignoring Max Score input...")
      GUIOutput.insert(END, total_score_error)
      output_list.append(total_score_error)
      output_list.append("")
      print(total_score_error)

  print("Processing Complete...")
  GUIOutput.insert(END,"")
  GUIOutput.insert(END,"Processing Complete...")
  output_list.append("")
  output_list.append("Processing Complete...")

  #Output to xlsx in path
  try:
    output_db.to_excel(Output_path+'\\Peer_Eval_output.xlsx')
    output_db_msg = ("Peer_Eval_output.xlsx saved in "+ Output_path)
    print(output_db_msg)
    GUIOutput.insert(END,output_db_msg)
    output_list.append(output_db_msg)
  except Exception as error:
    print("Error... Unable to output to excel file. Please close Peer_Eval_output.xlsx")
    print(error)
    GUIOutput.delete(0,END)
    GUIOutput.insert(END,"Error... Unable to output to excel file. Please close Peer_Eval_output.xlsx")
    GUIOutput.insert(END, error)
    output_list.append(error)

  #Generate reviewer name in excel wb as a comment
  if True:
    wb = openpyxl.load_workbook(Output_path+'\\Peer_Eval_output.xlsx')
    ws = wb.active    
    for x in range(len(score_comment_list)):
      cell_value = score_comment_list[x][0]
      name = score_comment_list[x][1]
      comment_cell = Comment(text= f'Reviewer: {name}', author = f'{name}')
      ws[f'{cell_value}'].comment = comment_cell
    wb.save(Output_path+'\\Peer_Eval_output.xlsx')

    
  #if txt output is selected, save as .txt file
  if output_txt.get() == TRUE:
    try:
      os.chdir(Output_path)
      txt_output = open(file= 'Output_log.txt',mode= 'w')
      for line in output_list:
        txt_output.write(f"{line}\n")
      txt_output.close
      txt_output_msg = ("Output_log.txt saved in "+Output_path)
      print(txt_output_msg)
      GUIOutput.insert(END,txt_output_msg)
      output_list.append(txt_output_msg)
    except Exception as error:
      print(error)
      GUIOutput.insert(END,error)
      output_list.append(error)

  #If non submitted users is selected, save as .txt file
  if output_missing_users.get() == TRUE:
    missing_user_list = []
    missing_user_list = db["login_id"].values.tolist()
    missing_user_list = list(set(missing_user_list) - set(UsernameList_track))
    try:
      os.chdir(Output_path)
      missing_output = open(file = 'Missing_users.txt', mode = 'w')
      for line in missing_user_list:
        missing_output.write(f"{line}\n")
      missing_output.close
      missing_output_msg = ("Missing_users.txt saved in "+ Output_path)
      missing_output_msg2 = ("There were "+ str(len(missing_user_list)) +" student(s) that did not submit a response.")
      print(missing_output_msg)
      GUIOutput.insert(END,missing_output_msg)
      output_list.append(missing_output_msg)
      print(missing_output_msg2)
      GUIOutput.insert(END,missing_output_msg2)
      output_list.append(missing_output_msg2)
    except Exception as error:
      print(error)
      GUIOutput.insert(END,error)
      output_list.append(error)
    
def close_it():
    Path_GF.insert(1,"Quitting Program...")
    Path_GF.delete(0,END)
    quit()

def browseFiles_db():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        Path_db.delete(0, END)
        Path_db.insert(END, file_path)



def browseFiles_GF():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        Path_GF.delete(0, END)
        Path_GF.insert(END, file_path)

def browseFiles_Output():
    file_path = filedialog.askdirectory()
    Path_Output.delete(0,END)
    Path_Output.insert(1,file_path)

# create root window
root = Tk()

# root window title and dimension
root.title("Peer Evaluation Score Calculator")
# Set geometry (widthxheight)
root.geometry('670x600')
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
GUIOutput.place(x=20,y=270,width=620,height=300)

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
Go_Button["text"] = "Run"
Go_Button.place(x=550,y= 205,width=70,height=50)
Inputdb_path = Path_db.get(1)
GoogleForms_path = Path_GF.get(1)
Go_Button["command"] = Run


Label4 = Label(root,text="Number of Questions:")
ft = tkFont.Font(family='MS Sans',size=10)
Label4["font"] = ft
Label4.place(x=30,y=165)

q_options = range(1,9)
# datatype of menu text
clicked1 = StringVar()
# initial menu text
clicked1.set( 6 )
# Create Dropdown menu
drop1 = OptionMenu( root , clicked1 , *q_options )
drop1.pack()
drop1.place(x=170, y=160)

Label5 = Label(root,text="Matching Probability Warning Threshold (%):")
ft = tkFont.Font(family='MS Sans',size=10)
Label5["font"] = ft
Label5.place(x=250,y=165)
q_options = range(1,9)
# datatype of menu text
clicked2 = StringVar()
# initial menu text
clicked2.set( 85 )
# Create Dropdown menu
matching_options = (80, 85, 90, 95)
drop2 = OptionMenu( root , clicked2 , *matching_options )
drop2.pack()
drop2.place(x=540, y=160)

Label6 = Label(root,text="Optional Settings:")
Label6["font"] = ft
Label6.place(x=30,y=200)

#Checkboxes
output_txt = IntVar()
output_missing_users = IntVar()

c1 = Checkbutton(root, text="Output_log.txt",variable= output_txt, onvalue= TRUE, offvalue= FALSE)
c1.place(x= 30, y= 230)
c2 = Checkbutton(root, text='Missing_users.txt',variable= output_missing_users, onvalue= TRUE, offvalue= FALSE)
c2.place(x= 160, y= 230)

#total score calculator
total_label = Label(root, text = "Max Score:")
total_label.place( x = 300, y= 230)
total_label_box = Entry(root)
total_label_box.place( x = 370, y = 230)

#icon
root.iconbitmap(r'C:\\Users\\adria\\OneDrive\\Desktop\\PeerEval\\logo.ico')
# Execute Tkinter
root.mainloop()
