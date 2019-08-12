# BEFORE RUNNING CODE, SAVE ALL .MSG FILES TO A SINGLE FOLDER

import os
#import fnmatch
import win32com.client #INSTALLED AS pywin32
import glob
import pandas as pd

# UPDATE THE FOLLOWING VARIABLES BEFORE RUNNING CODE
msg_folder = '..\\msg_folder' #UPDATE THIS PATH TO THE FOLDER WHERE .MSG ARE SAVED
att_folder = '..\\msg_folder\\attachments\\' #UPDATE THIS PATH TO THE FOLDER WHERE YOU WANT THE ATTACHMENTS SAVED
att_name_match = 'daily report*.xlsx' #UPDATE WITH FILE NAMES IN ATTACHMENT FOLDER (USE * TO DENOTE WILDCARD)
export_fname = '..\\msg_folder\\attachments\daily_reports_combined.csv' #UPDATE WITH PATH AND FINAL FILE NAME
#

file_list = []

# LOAD ALL .MSG FILE NAMES INTO A LIST
for filename in os.listdir(msg_folder): 
	# IF YOU NEED TO MATCH SPECIFIC FILE NAME PATTERNS, UNCOMMENT AND UPDATE THE LINE BELOW AND INDENT THE APPEND STATEMENT
    #if fnmatch.fnmatch(filename, '*DOC_PART*.csv') and not fnmatch.fnmatch(filename, '*CLEAN*'):
    if filename.find('.msg') > 0:
        file_list.append(os.path.join(msg_folder, filename))

# EXTRACT ATTACHMENT FROM EACH .MSG FILE IN FILE_LIST
counter = 1
for file in file_list:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(file)
    att = msg.Attachments
    sub_count = 0 #USED FOR SINGLE EMAILS WITH MULTIPLE ATTACHMENTS
    for a in att:
        fname = (a.FileName[:a.FileName.find('.xlsx')] + '%s%s' + '.xlsx') % (counter, sub_count)
        a.SaveAsFile(os.path.join(att_folder, fname))
        sub_count += 1
    counter += 1
    del outlook, msg

# START MERGE EXCEL FILES
all_data = pd.DataFrame()
total_len = 0
for f in glob.glob(att_folder+att_name_match):
    df = pd.read_excel(f, skiprows=[0]) # skiprows=[0] SKIPS THE FIRST ROW IN THE EXCEL DOC - MY REPORT HAD A HEADER ABOVE THE COLUMN HEADERS I WANTED TO IGNORE
    all_data = all_data.append(df,ignore_index = True)

# REMOVE ROWS WHERE ALL COLUMNS ARE NULL
print('starting num rows:', len(all_data))
all_data = all_data.dropna(how='all')
print('num rows after dropping null:', len(all_data))

# EXPORT DATA TO CSV FOR UPLOAD INTO DATABASE OR TABLEAU USE
# FILE EXPORTED AS TAB SEPARATED AS THIS CAUSES FEWER ISSUES WHEN UPLOADING VIA TOAD DATA POINT
all_data.to_csv(export_fname, index = False, sep='\t')
