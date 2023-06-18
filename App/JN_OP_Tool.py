#!/usr/bin/env python
# coding: utf-8

# In[1]:


# // Dependencies
print("Loading Dependencies\n--------------------")

import warnings
warnings.catch_warnings()
warnings.simplefilter("ignore")
import pandas as pd
print("10%")
import matplotlib
print("20%")
matplotlib.use('TKAgg')
print("30%")
import matplotlib.pyplot as plt
print("40%")
import sys
print("50%")
import os
print("60%")
import tkinter as tk
print("70%")
from tkinter import ttk
print("80%")
from tkinter import messagebox
print("90%")
from tkinter import filedialog
import xlsxwriter
print("100%")
import glob
print("--------------------\nLoaded")


# In[2]:


# // Path to Query
#folder_path = 'QUERY_FILE_GOES_HERE/'

#xlsx_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

#QUERYS = {}
#Q_Index = 0

#for file_path in xlsx_files:
#    OPName = os.path.splitext(os.path.basename(file_path))[0] # Extract filename without extension as DataFrame name
#    QUERYS[OPName] = pd.read_excel(file_path)
#    Q_Index += 1
    
#print(f"Files in folder '{folder_path}':")

#if Q_Index > 0:
#    for file_name in QUERYS.keys():
#        print(file_name)
#else:
#   print(f"No .xlsx file found in {folder_path}")


# In[3]:


root = tk.Tk()
Query_Path = filedialog.askopenfilename(initialdir='QUERY_FILE_GOES_HERE/', title='Select a file', filetypes=[('Excel files', '*.xlsx')])
Query = pd.read_excel(Query_Path)
terms = []
enroll_status_cols = Query.filter(like='_Enroll_Status').columns
for term in enroll_status_cols:
    term = term[:6]
    terms.append(term)
    print(term)


# In[4]:


root.destroy()
root = tk.Tk()
root.title("Current Term Input")

label = tk.Label(root, text="Input current term in Banner format: \nOptions are:")
label.pack()
label = tk.Label(root, text="\n".join(terms))
label.pack()

entry = tk.Entry(root)
entry.pack()

def get_term():
    global Current_Term
    global LastTwo
    Current_Term = entry.get()
    if Current_Term.isdigit():
        Current_Term = int(Current_Term)
        LastTwo = Current_Term % 100
        if Current_Term > 100000 and Current_Term < 999999 and (LastTwo == 10 or LastTwo == 20 or LastTwo == 30):
            root.destroy()
        else:
            messagebox.showerror("Error", "Input not in correct Banner format")
    else:
        messagebox.showerror("Error", "Input is not a number")

submit_btn = tk.Button(root, text="Submit", command=get_term)
submit_btn.pack()

root.mainloop()
print(f"Term selected is " + str(Current_Term))


# In[5]:


# Returns Count of VIDs as Population
TPS = Query["VID"].count()

# Sets Term Codes & Reads Current Term from input
if LastTwo == 10:
#    Fall = Current_Term
    Term_Actual = "Fall"
elif LastTwo == 20:
#    Fall = Current_Term - 10
    Term_Actual = "Spring"
elif LastTwo == 30:
#    Fall = Current_Term - 20
    Term_Actual = "Summer"
#    Spring = Fall + 10
#    Summer = Spring + 10

if LastTwo == 30:
    NTerm_Search = str(Current_Term + 80) + "_Enroll_Status" # Checks to see if Query Term is Summer and adjusts search
    ANTerm_Search = str(Current_Term + 90) + "_Enroll_Status"
elif LastTwo == 20:
    NTerm_Search = str(Current_Term + 10) + "_Enroll_Status"
    ANTerm_Search = str(Current_Term + 80) + "_Enroll_Status"
elif LastTwo == 10:
    NTerm_Search = str(Current_Term + 10) + "_Enroll_Status"
    ANTerm_Search = str(Current_Term + 20) + "_Enroll_Status"

CTerm_Search = str(Current_Term) + "_Enroll_Status" # Column name for query term
CTerm = Query[CTerm_Search].count() 
Current_E_List = Query[Query[CTerm_Search].notnull()] # Students In Query Term
Current_Not_EL = len(Current_E_List[Current_E_List[CTerm_Search] != "EL"]) # Students In Query Term Withdrawn, Dropped, or Other
Current_WD = len(Current_E_List[Current_E_List[CTerm_Search] == "WT"]) # Students In Query Term Withdrawn
    
NTerm = Query[NTerm_Search].count()
Next_E_List = Query[Query[NTerm_Search].notnull()] # Students In Next Term
Next_Not_EL = len(Next_E_List[Next_E_List[NTerm_Search] != "EL"]) # Students In Next Term Withdrawn, Dropped, or Other
Next_WD = len(Next_E_List[Next_E_List[NTerm_Search] == "WT"]) # Students In Next Term Withdrawn

if ANTerm_Search in enroll_status_cols:
    ANTerm = Query[ANTerm_Search].count()
    AfterNext_E_List = Query[Query[ANTerm_Search].notnull()] # Students In 2 Terms out
    AfterNext_Not_EL = len(AfterNext_E_List[AfterNext_E_List[ANTerm_Search] != "EL"]) # Students In 2 Terms out Withdrawn, Dropped, or Other
    AfterNext_WD = len(AfterNext_E_List[AfterNext_E_List[ANTerm_Search] == "WT"]) # Students In 2 Terms out Withdrawn

#DA_NTerm_List = Query[Query["TermCodeAdmit"] > Current_Term]
#Num_DA_NTerm = DA_NTerm_List[NTerm_Search].count()

# Index for graphs: count of all students admitted during or after next term
#DASIndex = len(Query[Query["TermCodeAdmit"] > Current_Term].count())


# 1. Total registered in all available terms
# 2. From next term -> # registered in after next term (Case: From fall -> spring)
# 3. From admitted current term -> registered for next term (Case: from summer -> fall)
# 4. From admitted current term -> registered in after next term (Case: from summer -> spring)
# 5. From admitted after next term -> registered (Case: from spring admitted -> spring enrolled)
# 6. From current term -> withdrew 
# 7. Numbers for methods of instruction for each term // Pending
# 8. Total numbers per term, Full-Time & Part-Time
# 9. Export CSV - Students who withdrew or not registered

# In[6]:


# 2. From Next Term -> Enrolled After Next Term
if ANTerm_Search in enroll_status_cols:
    Next_AfterNext_EL = len(Next_E_List[Next_E_List[ANTerm_Search] == "EL"]) #counts enrolled in two terms after current from next term registered list
    NtoANratio = Next_AfterNext_EL / NTerm

    labels = ['Enrolled\n' + str(Next_AfterNext_EL) + ' Students',
              'Not Enrolled\n' + str(NTerm - Next_AfterNext_EL) + ' Students'
             ]
    sizes = [NtoANratio, 1-NtoANratio]

    fig, ax = plt.subplots()
    ax.pie(sizes, labels=labels, startangle=90, counterclock=False, autopct='%1.1f%%')

    plt.title('Students Who Enrolled in ' + str(NTerm_Search[0:6]) + '\n That Enrolled in ' + str(ANTerm_Search[0:6]))

    plt.show()


# In[7]:


# 3. From admitted current term -> registered for next term (Case: from summer -> fall)
CTerm_Admits = Query[Query['TermCodeAdmit'] == Current_Term]
N_EN_CTerm_Admits = CTerm_Admits[CTerm_Admits[NTerm_Search].notnull()]
R3Ratio = len(N_EN_CTerm_Admits) / len(CTerm_Admits)

labels = ['Enrolled\n' + str(len(N_EN_CTerm_Admits)) + ' Students',
          'Not Enrolled\n' + str(len(CTerm_Admits) - len(N_EN_CTerm_Admits)) + ' Students'
         ]
sizes = [R3Ratio, 1-R3Ratio]

fig, ax = plt.subplots()
ax.pie(sizes, labels=labels, startangle=90, counterclock=False, autopct='%1.1f%%')

plt.title('Students Who Were Admitted in ' + str(Current_Term) + '\n That Enrolled in the ' + str(NTerm_Search[0:6]))

plt.show()


# In[8]:


# 4. From admitted current term -> registered in after next term (Case: from summer -> spring)
R4Search = int(NTerm_Search[0:6])
NTerm_Admits = Query[Query['TermCodeAdmit'] == R4Search]
if ANTerm_Search in enroll_status_cols:
    AN_EN_CTerm_Admits = CTerm_Admits[CTerm_Admits[ANTerm_Search].notnull()]
    R4Ratio = len(AN_EN_CTerm_Admits) / len(CTerm_Admits)

    labels = ['Enrolled\n' + str(len(AN_EN_CTerm_Admits)) + ' Students',
              'Not Enrolled\n' + str(len(CTerm_Admits) - len(AN_EN_CTerm_Admits)) + ' Students'
             ]
    sizes = [R4Ratio, 1-R4Ratio]

    fig, ax = plt.subplots()
    ax.pie(sizes, labels=labels, startangle=90, counterclock=False, autopct='%1.1f%%')

    plt.title('Students Who Were Admitted in ' + str(Current_Term) + '\n That Enrolled in the ' + str(ANTerm_Search[0:6]))

    plt.show()


# In[9]:


# 5. From admitted after next term -> registered (Case: from spring admitted -> spring enrolled)
if ANTerm_Search in enroll_status_cols:
    R5Search = int(ANTerm_Search[0:6])
    ANTerm_Admits = Query[Query['TermCodeAdmit'] == R5Search]
    AN_EN_ANTerm_Admits = ANTerm_Admits[ANTerm_Admits[ANTerm_Search].notnull()]

    R5Ratio = len(AN_EN_ANTerm_Admits) / len(ANTerm_Admits)

    labels = ['Enrolled\n' + str(len(AN_EN_ANTerm_Admits)) + ' Students',
          'Not Enrolled\n' + str(len(ANTerm_Admits) - len(AN_EN_ANTerm_Admits)) + ' Students'
         ]
    sizes = [R5Ratio, 1-R5Ratio]

    fig, ax = plt.subplots()
    ax.pie(sizes, labels=labels, startangle=90, counterclock=False, autopct='%1.1f%%')

    plt.title('Students Who Were Admitted in ' + ANTerm_Search[0:6] + '\n That Enrolled in the ' + str(ANTerm_Search[0:6]))

    plt.show()


# In[10]:


# 6. From current term -> withdrew
WDRatio = Current_WD / CTerm
Other_Ratio = (Current_Not_EL - Current_WD) / CTerm

labels = [
        'Withdrawn\n' + str(Current_WD) + ' Students', 
          'Other\n' + str(Current_Not_EL - Current_WD) + ' Students', 
          'Enrolled\n' + str(CTerm - (Current_Not_EL - Current_WD)) + ' Students'
         ]

sizes = [WDRatio, Other_Ratio, 1-WDRatio-Other_Ratio]

fig, ax = plt.subplots()
ax.pie(sizes, labels=labels, startangle=180, counterclock=False, autopct='%1.1f%%')

plt.title('Students Enrolled in ' + str(Current_Term) + '\nThat Withdrew')

plt.show()


# In[11]:


# 8. Total numbers per term, Full-Time & Part-Time
Load_df = pd.DataFrame()
for term in terms:
    Load_Search = "Enrolled_" + str(term)
    Fulltime = Query[Load_Search][Query[Load_Search] >= 12].count()
    Parttime = Query[Load_Search][(Query[Load_Search] >= 6) & (Query[Load_Search] < 12)].count()
    LessParttime = Query[Load_Search][Query[Load_Search] < 6].count()
    Load_df = Load_df.append({'Term':term, 'FT':Fulltime, 'PT':Parttime, 'Less Than PT':LessParttime}, ignore_index=True)
print(Load_df)


# In[12]:


# 9. Export CSV - Students who withdrew or not registered
writer = pd.ExcelWriter('Student Withdrawls & Non-Enrollees.xlsx', engine='xlsxwriter')

CT_S_WD = Current_E_List[Current_E_List[CTerm_Search] == "WT"]
CT_S_WD.to_excel(writer, sheet_name=str(CTerm_Search[0:6]) + ' Withdrawls', index=False)
NT_S_WD = Next_E_List[Next_E_List[NTerm_Search] == "WT"]
NT_S_WD.to_excel(writer, sheet_name=str(NTerm_Search[0:6]) + ' Withdrawls', index=False)
if ANTerm_Search in enroll_status_cols:
    ANT_S_WD = AfterNext_E_List[AfterNext_E_List[ANTerm_Search] == "WT"]
    ANT_S_WD.to_excel(writer, sheet_name=str(ANTerm_Search[0:6]) + ' Withdrawls', index=False)

CT_S_Nenroll = CTerm_Admits[CTerm_Admits[CTerm_Search].isnull()]
CT_S_Nenroll.to_excel(writer, sheet_name='Admitted & Not Enrolled ' + str(CTerm_Search[0:6]), index=False)
NT_S_Nenroll = NTerm_Admits[NTerm_Admits[NTerm_Search].isnull()]
NT_S_Nenroll.to_excel(writer, sheet_name='Admitted & Not Enrolled ' + str(NTerm_Search[0:6]), index=False)
if ANTerm_Search in enroll_status_cols:
    ANT_S_Nenroll = ANTerm_Admits[ANTerm_Admits[ANTerm_Search].isnull()]
    ANT_S_Nenroll.to_excel(writer, sheet_name='Admitted & Not Enrolled ' + str(ANTerm_Search[0:6]), index=False)

writer.save()


# In[13]:


#CTerm_Admit = Query[Query['']]
#NTerm_Admit
TCAIndex = Query['TermCodeAdmit']
TCAList = Query['TermCodeAdmit'].unique()
Filtered_TCAList = []
for i in TCAList:
    if i > 0:
        Filtered_TCAList.append(i)
Filtered_TCAList.sort()
print(Filtered_TCAList)


# In[14]:


with open("report.txt", "w") as file:
    file.write("Osceola Prosper Report: Reference Term " + str(Current_Term) + 
               "\n------------------------------------------")
# 1. Total registered in all available terms    
    file.write("\nTotal Prosper Students: " + str(TPS) + 
               "\n\nENROLLMENT BY TERM\n- - -\n" + str(Current_Term) + " Enrolled Students: " + str(CTerm) + 
               "\n" + str(NTerm_Search[0:6]) + " Enrolled Students: " + str(NTerm))
    if ANTerm_Search in enroll_status_cols:
        file.write("\n" + str(ANTerm_Search[0:6]) + " Enrolled Students: " + str(ANTerm))
        
    file.write("\n\nADMITS BY TERM\n- - -\n")
    for i in Filtered_TCAList:
        file.write(str(i)[:6] + " Admitted Students: " + str(len(TCAIndex[TCAIndex==i])) + "\n")
    file.write("- - -")

# 2. From next term -> # registered in after next term (Case: From fall -> spring)
    if ANTerm_Search in enroll_status_cols:
        file.write("\n\nOf Students Enrolled in " + str(NTerm_Search[0:6]) + " " + "(" + str(NTerm) + ")" + " | " + 
                   str(Next_AfterNext_EL) + " Students are Enrolled in " + str(ANTerm_Search[0:6]))

    # 3. From admitted current term -> registered for next term (Case: from summer -> fall)
    file.write("\n\nOf Students Admitted in " + str(Current_Term) + " " + "(" + str(len(CTerm_Admits)) + ")" + " | " + 
               str(len(N_EN_CTerm_Admits)) + " Students are Enrolled in " + str(NTerm_Search[0:6]))
# 4. From admitted current term -> registered in after next term (Case: from summer -> spring)
    if ANTerm_Search in enroll_status_cols:
        file.write("\nOf Students Admitted in " + str(Current_Term) + " " + "(" + str(len(CTerm_Admits)) + ")" + " | " + 
                   str(len(AN_EN_CTerm_Admits)) + " Students are Enrolled in " + str(ANTerm_Search[0:6]))
# 5. From admitted after next term -> registered (Case: from spring admitted -> spring enrolled)
        file.write("\nOf Students Admitted in " + str(ANTerm_Search[0:6]) + " " + "(" + str(len(ANTerm_Admits)) + ")" + " | " + 
                   str(len(AN_EN_ANTerm_Admits)) + " Students are Enrolled")

# 6. From current term -> withdrew
    file.write("\n\nOf Students Enrolled in " + str(Current_Term) + " " + "(" + str(len(CTerm_Admits)) + ")" + " | " + 
              str(Current_WD) + " Students Withdrew\n\n")
    
# 7. Numbers for methods of instruction for each term // Pending

# 8. Total numbers per term, Full-Time & Part-Time
    file.write("------------------------------------------\nNumber of Students per Term & Credit Workloads\n\n" + 
               str(Load_df) + "\n\nSee Student Withdrawls & Non-Enrollees.xlsx for List of Student Withdrawls & Non-Enrollment Per Semester.")


# In[ ]:




