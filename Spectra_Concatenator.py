#If you've never used python before, run the other .py file in the repo first to install these libraries. 
import glob
import pandas as pd
import os
os.chdir(
        #Path of spectra in .csv format. 
        r"E:\Blake-Arthur\MgO-Support-(O-R-400C)E-(R-200)I\1-0 Spectra taken of In-situ Reduction\0-6-CS(S-N2-20C)-BG(EH-N2-20C)-PostRed"
)

#Sort based on #'s contained in file name. If you have an error at the indicated line, you have .csv file(s) in the target folder whos name does not contain numbers. 
files = glob.glob('*.CSV')
files.sort(key=lambda x: int("".join([i for i in x if i.isdigit()]))) # error here caused by explanation above. if not, troubleshoot by print(files) on the line above this. 
print(files)

#Assign a unique, evenly-numbered, name to wavenumber files to identify unnecessary columns for deletion.
#If you delete a column referenced by position, and it shares a name with another column, both columns will be deleted regardless of their positions.   
combo_table = pd.concat((pd.read_csv(f,skip_blank_lines=True,header=None,names=[files.index(f)*2,"Trial "+str(files.index(f))]) for f in files), axis=1) 
#Find unnecessary Wavenumber columns.
length = list(range(len(combo_table.columns)))
list = []
for i in length:
    if i//2 >= 1 and i%2 == 0:
        list.append(i)          
    else:
        pass
#Delete unnecessary Wavenumber columns.
proc_table = combo_table.drop(combo_table.columns[list],axis=1)

print("\n---------------------------------------------------------------\n\nPreview:\n", proc_table)

#Use foldername as filename, write to .xlsx on active directory. 
path = os.getcwd().split("\\")
foldername = (path[len(path)-1])
proc_table.to_excel(foldername+".xlsx")
