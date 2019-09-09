
import os
from tkinter import messagebox
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# ... DATA FILTERING
def work(df,...,output,old_excel):

    #reading the excels because of the possibility that they changed and didnt selected again.
    try:
        df = pd.read_excel(...,index_col=None)
    except:
        messagebox.showerror("ERROR","Please close ... Excel file and select again!")
        return
    try:
        ... = pd.read_excel(...,index_col=None)
    except:
        messagebox.showerror("ERROR","Please Close ... Excel file and select again!")
        return
    try:
        old_excel = pd.read_excel(output,index_col=None)
    except:        
        messagebox.showerror("ERROR","Please close output.xlsx and output_empty....xlsx files and select again!")
        return


    #initialize error array 
    error_array = []

    #filtering the datas that have 0 at ... column
    try:
        df = df[df['...'] != 0]
    except:
       # messagebox.showerror("ERROR:", "")
        error_array.append("... Excel does not have any Dlv.qty column. Please fix it.")
    
    #filtering the datas that does not contain ... at ...
    try:
        df = df[df['...'].str.contains("...")]
    except:
        error_array.append("... Excel does not have any ... column. Please fix it.")

    #getting ... column  as pandas.core.series.Series
    try:
        billoflading = df['...']
    except:
        error_array.append("... Excel does not have any ... column. Please fix it.")

    #error control for ... Columns
    if len(error_array) != 0:
        messagebox.showerror("... COLUMN ERROR",error_array)
        return

    #splitting the data and creating a stack
    ... = df['...'].str.split(',')

    #creating arrays to getting split values
    ... = []
    ... = []
    ... = []
    ... = []
    ... = []
    #loop to get every each row of ...
    for x in ...:

        #if we have no data, we initialize None
        if(not isinstance(x,list)):
            ....append(None)
            ....append(None)
            ....append(None)
            ....append(None)
            ....append(None)
            continue

        # if we have not requered amount of data, we initalize to None
        if(len(x)<3):
            ....append(None)
            ....append(None)
            ....append(None)
            ....append(None)
            ....append(None)
            continue 

        #we have desired amount of data
        else:

            #flags for detecting any absence of one them
            flagL = False
            flagC = False
            flagI = False
            flagT = False
            #counts for detecting the order of them
            count_L = 0
            count_C = 0
            count_I = 0
            count_T = 0
            index = 0

            #for each partition of row
            for i in x:

                index += 1
                #removing any blank space that can be due to the manuel input mistakes
                if i.startswith(" "):
                    i = i.replace(" ","")

                #finding ...
                if "..." in i:
                    flagI = True
                    count_I += index
                    ....append(i)
                    continue

                #finding ...
                if i.startswith('...'):
                    flagT = True
                    count_T += index
                    ....append(i)
                    continue

                #finding ...
                elif i.startswith('0') or i.startswith('1') or i.startswith('2') or i.startswith('3') or i.startswith('4') or i.startswith('5') or i.startswith('6') or i.startswith('7') or i.startswith('8') or i.startswith('9'):
                    flagL = True
                    count_L += index
                    ....append(i)
                    continue

                #finding ...
                else:                
                    flagC = True
                    count_C += index
                    ....append(i)
                    continue

            #if flag is False that means it is absence so we are filling with None
            if flagI is False:
                ....append(None)
            if flagL is False:
                ....append(None)
            if flagC is False:
                ....append(None)
            if flagT is False:
                ....append(None)

            #Checking if the ... order is accommodate to our desire and stacking the information.
            if not ( count_L == 1 and count_C == 2 and count_I == 3 and count_T == 4):
                ....append("FALSE")
            else:
                ....append("TRUE")

    #adding our arrays as columns to dataframe     
    df['...'] = ...
    df['...'] = ...
    df['...'] = ...
    df['...'] = ...
    df['...'] = ...

    #filtering what is not None
    df = df[df.....notnull()]

    #deleting the original column
    del df['...']
    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    # MERGING DATAS 

    #creating a ...  that includes these columns from ...  and ... 
    try:
        ... = df[['...','...','...','....','...','...','...','...','...']]
    except:
        messagebox.showerror("... COLUMN ERROR","... Excel doesn't have one or more of these columns: \n ... \n ... \n ... \n ... \n ...")
        return
    try:
        ... = ...[['...', '...','...','...']]
    except:
        messagebox.showerror("... COLUMN ERROR", "... Excel doesn't have one or more of these columns: \n... \n ... \n ... \n ... ")
        return        

    # we merged ... and ... to one in order to get the uniqueness
    try:
        ...['Merged'] = sevk_rapor[['...', '...']].apply(lambda x: ','.join(x), axis=1)
    except:
        error_array.append("Couldn't Merge ... and ... Columns in ... Excel. This error could be because of wrong data type in these cells.")

    #we list the ... under same Merged 
    try:
        ...['...'] = sevk_rapor['...'].astype(str).str[:-2]
    except:
        error_array.append("... column in ... couldn't converted into a string value. There could be some comma(,) or dot(.) in these cells. ")

    try:
        ... = ....groupby(['Merged'])['...'].apply(lambda x: ','.join(x)).reset_index()
    except:
        error_array.append("Couldn't group ... column in ... by Merged(...).")

    #we sum these ...'s ... and change it
    try:
        ... = ....groupby(['Merged'])['...'].apply(lambda x:x.sum()).reset_index()
    except:
        error_array.append("Couldn't sum the ... in ... Excel for each Merged(...)")

    #we merged what we did into a new dataframe
    try:
        ... = ....merge(..., on = 'Merged')
    except:
        error_array.append("Couldn't merge the dataframes for .... Couldn't initialize ...")

    #delete the possible duplicate columns when we merged new dataframe and the old one
    try:
        del sevk_rapor['...']
        del sevk_rapor['...']
    except:
        error_array.append("Couldn't erase the Dlv.qty or ... in .... So there could be a duplicate rows of them.")

    #merged new and old 
    try:
        ... = ....merge(..., on = 'Merged')
    except:
        error_array.append("Couldn't merge two dataframe of .... There could be less rows due to this.")

    #error control for ... ERRORs
    if len(error_array) != 0:
        messagebox.showerror("... EXCEL DATA PROCESS ERRORS",error_array)
        return

    #sometimes when we convert ... into str '.0' comes with it. If it is happen, we remove the dot point to have a clean merged data
    flag_dot_zero = False   
    try:
        ...['...'] = ...['...'].astype(str)
    except:
       error_array.append("Couldnt convert ... in ... into string. Please check the data type. ")

    try:
        for x in  ...['...']:  
            if  '.' in x:
                flag_dot_zero = True

        if(flag_dot_zero == True):
            ...['...'] = ...['...'].astype(str).str[:-2]
    except:
        error_array.append("Couldn't remove the dot points in ... column ....")
    
    #merged ... and ... to have an Merged uniqueness
    try:
        ...['...'] = ...['...'].astype(str)
        ...['Merged']  = ...[['...', '...']].apply(lambda x: ','.join(x), axis=1)
    except:
        error_array.append("Couldnt merge ... and ... in ... Excel")

    #delete the possible duplicate columns when we merged new dataframe and ...
    try:
        del ...['...']
    except:
        error_array.append("Couldn't delete ... Column in ... Excel")

    #merged that into the 
    try:
        ... = ....merge(..., on ='Merged')
    except:
        error_array.append("Couldnt merge ... and ... together.")

    #error control for ... Errors
    if len(error_array) != 0:
        messagebox.showerror("... ERROR",error_array)
        return

    #to control if there is any rows that does not match and left at the .... That means their indiviual ... doesnt created yet.    
    ... = []
    try:
        for x in ...['Merged']:
            flag_exists = False
            for y in ...['Merged']:
                if x in y:
                    flag_exists = True

            if flag_exists is True:
                ....append("... VAR")

            else:
                ....append("... YOK")
        #making ... column
        ...['...'] = ...
        ... = ....loc[ (...['...'] == '...') & (...['...'] != "nan") ]
    except:
        error_array.append("...'s in ... couldnt find,match, etc. in ...'s ...")

    #making new excel that contains the rows that does not match
    try:
        ....to_excel(output + "_empty....xlsx")
    except:
        messagebox.showerror("ERROR","Please close output_empty....xlsx file.")
        return

    #adding control column to do the logicial desicion that if ... and ... are the same
    try:
        ... = []
        index_m = 0
 
        for x in ...['...']:
            index_m += 1
            index_d = 0
            
            for i in ...['...']:
                index_d += 1
                if(index_m == index_d):
                    if i == x:
                        ....append("TRUE")

                    else:
                        ....append("FALSE")

        ...['...'] = control_adet
    except:
        error_array.append("... Column couldnt created.")

    #differentiate append and old data
    try:
        ...['...'] = "..."
    except:
        error_array.append("Couldnt make ... Column.")

    #merge the new data with preious excel 
    try:
        if old_excel.empty is not True:
            #differentiate append and old data
            old_excel['...'] = '...'
            old_excel.reset_index(drop=True)
            ... = old_excel.append(...,ignore_index= True, sort = False )
    except:
        error_array.append("Couldnt write on the output excel. Please look at the output excel.")

    #dropping duplicate rows
    try:
        ... =  ....drop_duplicates(keep= False)
    except:
        error_array.append("Couldnt drop the duplicate rows.")

    if len(error_array) != 0:
        messagebox.showerror("ERRORS",error_array)

    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    # HIGHLIGHTING THE COLUMNS

    #writing our DataFrame into an excel sheet called 'Sayfa1'
    try:
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        ....to_excel(writer, sheet_name='Sheet1',index=False)
        workbook  = writer.book

        #initializing our highlight colors for excel rows
        yellow = workbook.add_format({'bg_color': 'yellow'})
        red = workbook.add_format({'bg_color': 'red'})
        green = workbook.add_format({'bg_color': 'green'})
        worksheet = writer.sheets['Sheet1']

        #highlighting ... Column
        worksheet.conditional_format('N2:N500000', {'type': 'text',
                                                'criteria': 'containing',
                                                'value': 'TRUE',
                                                'format': green})

        worksheet.conditional_format('N2:N500000', {'type': 'text',
                                                'criteria': 'containing',
                                                'value': 'FALSE',
                                                'format': red})
        #Highlighting ... Column
        worksheet.conditional_format('J2:J500000', {'type': 'text',
                                                'criteria': 'containing',
                                                'value': 'TRUE',
                                                'format': green})

        worksheet.conditional_format('J2:J500000', {'type': 'text',
                                                'criteria': 'containing',
                                                'value': 'FALSE',
                                                'format': red})

        #Highlightind ... Column
        worksheet.conditional_format('O2:O500000', {'type': 'text',
                                                'criteria': 'containing',
                                                'value': 'NEW',
                                                'format': yellow})
                            
        #saving what we did on excel
        writer.save()

        #opening the both output excels after execution
        os.startfile(output)
        os.startfile(output + "_empty....xlsx")
    except:
        error_array.append("Couldnt highlight the rows")
    if len(error_array) != 0:
        messagebox.showerror("ERROR",error_array)
        return
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# GUI MAKING

#making root 
root = tk.Tk(className=' Compare_Excels')
canvas1 = tk.Canvas(root, width=600, height=450, bg = 'lightsteelblue')
canvas1.pack()

#Making Entryies for each button to display selected file names
... = tk.Entry(canvas1, width=40, bg='white', fg='black')
....configure(state='readonly')
....place(rely= 0.20, relx= 0.03)

... = tk.Entry(canvas1, width=40, bg='white', fg='black')
....configure(state='readonly')
....place(rely= 0.20, relx= 0.57)

OUTPUT = tk.Entry(canvas1, width=40, bg='white', fg='black')
OUTPUT.configure(state='readonly')
OUTPUT.place(rely= 0.43, relx= 0.325)

#... excel initialization
def get...Excel():
    global df
    global ...
    global ...
    ... = filedialog.askopenfilename()
    ... = os.path.basename(...)

    #updating entry box with the file name 
    ....configure(state='normal')
    ....delete(0, tk.END)
    ....insert('end', ...)
    ....configure(state='readonly')
        
    if ... == "":
        messagebox.showerror("ERROR", "Please Select ...  Excel")
    else:
        if (".xlsx" in ...) is True:
            try:
                df = pd.read_excel(...,index_col=None)
            except:
                messagebox.showerror("ERROR","Please close ... Excel file.")
        else:
            messagebox.showerror("ERROR", "Please Select a .XLSX form ...  Excel")

#... Excel initialization
def get...Excel():
    global ...
    global ......
    global ...
    ... = filedialog.askopenfilename()
    ... = os.path.basename(...)

    #updating entry box with the file name 
    ....configure(state='normal')
    ....delete(0, tk.END)
    ....insert('end', ...)
    ....configure(state='readonly')

    if ... == "":
        messagebox.showerror("ERROR", "Please Select ...  Excel")
    else:
        if (".xlsx" in ...) is True:
            try:
                ... = pd.read_excel(v,index_col=None)
            except:
                messagebox.showerror("ERROR","Please Close ... Excel file")
        else:
            messagebox.showerror("ERROR", "Please Select a .XLSX format ...  Excel")

#Output file location initialization
def getOutput():
    global output,old_excel
    global filename_output
    global import_file_path_output

    import_file_path_output = filedialog.askopenfilename()
    filename_output = os.path.basename(import_file_path_output)

    #updating entry box with the file name 
    OUTPUT.configure(state='normal')
    OUTPUT.delete(0, tk.END)
    OUTPUT.insert('end', filename_output)
    OUTPUT.configure(state='readonly')

    if import_file_path_output == "":
        messagebox.showerror("ERROR", "Please Select Output Location")
    output = import_file_path_output
    try:
        old_excel = pd.read_excel(output,index_col=None)
    except:
        messagebox.showerror("ERROR","Please close output.xlsx and output_empty....xlsx files")

#start button function that calls work function
def start():
    work(df,...,output,old_excel) 

#Making Buttons
browseButton_Excel... = tk.Button( text = 'Import ... Excel File',command=get...Excel, bg = 'green', fg='white', font=('helvetiva',12,'bold'))
browseButton_Excel... = tk.Button( text = 'Import ... Excel File',command=get...Excel, bg = 'green', fg='white', font=('helvetiva',12,'bold'))
browseButton_Output = tk.Button( text = 'Select Desired Output File',command=getOutput, bg = 'green', fg='white', font=('helvetiva',12,'bold'))
browseButton_Start = tk.Button( text = 'Start',command=start, bg = 'red', fg='white', font=('helvetiva',12,'bold'))

#Bounding the canvas
canvas1.create_window(150,150,window=browseButton_Excel...)
canvas1.create_window(450,150,window=browseButton_Excel...)
canvas1.create_window(300,250,window=browseButton_Output)
canvas1.create_window(300,350,window=browseButton_Start)

root.mainloop()

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

