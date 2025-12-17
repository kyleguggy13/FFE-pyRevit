__title__     = "Pressure Loss \nReport Reader"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 12.17.2025
_________________________________________________________________
Description:
Read Pressure Loss Report HTML file.

_________________________________________________________________
How-to:

_________________________________________________________________
Last update:
- [12.17.2025] - v0.1 BETA RELEASE
_________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import forms
from pyrevit.script import output

#____________________________________________________________________ IMPORTS (CUSTOM)
# from HTMLparse import HTMLparse as HTMLparse
from HTMLParser import HTMLParser


#____________________________________________________________________ VARIABLES
output_window = output.get_output()

log_status = ""
action = "Pressure Loss Report Reader"


#____________________________________________________________________ MAIN

# Pick HTML File
source_file = forms.pick_file(file_ext='HTML')

output_window.print_md(source_file)



# from bs4 import BeautifulSoup
# import os
import codecs

# with open(source_file, 'rt') as f:
#     html_content = f.read()

with codecs.open(source_file, "r", "utf-8") as f:
    html_text = f.read()

# soup = BeautifulSoup(html_content, 'html5lib')

output_window.print_md(html_text)



"""
### Assign values from GUI ###
Input_ImportPath = values['Input_ImportPath']
Input_excelpath = values['Input_excelpath']
Input_ImportCP = values['Input_ImportCP']
Input_sheetname_All = values['Input_sheetname_All']
Input_sheetname_CP = values['Input_sheetname_CP']


### Match strings in HTML ###
match1 = 'Detail Information of Straight Segment by Sections'
match2 = 'Fitting and Accessory Loss Coefficient Summary by Sections'
match3 = 'Total Pressure Loss Calculations by Sections'

### Read HTLML file || Create DataFrames ###
list_tables = pd.read_html(Input_ImportPath, flavor='html5lib')
list_tables_CP = pd.read_html(Input_ImportPath, match=match3, flavor='html5lib', header=2)
list_tables_Duct = pd.read_html(Input_ImportPath, match=match1, flavor='html5lib', header=2)
list_tables_Fittings = pd.read_html(Input_ImportPath, match=match2, flavor='html5lib', header=2)




#____________________________________________________________________ FUNCTIONS
def GetCriticalPaths(HTMLtable):

    CriticalPath = HTMLtable.iloc[-1,0].split(' ')[3].split('-')

    return CriticalPath


### Get Critical Path from report ###
list_CriticalPaths = []
for t in list_tables_CP:
    val = GetCriticalPaths(t)
    list_CriticalPaths.append(val)



#######################################
### Empty list
list_System_DuctData = []

### Send HTML tables to HTMLparse.py for parsing
### Parse HTML file for critical paths, duct data, & fittings data 
### Append DataFrames to list_System_DuctData
for CP, Duct, Fittings in zip(list_tables_CP, list_tables_Duct, list_tables_Fittings):
    list_dfData = HTMLparse(table_CP=CP, table_Duct=Duct, table_Fittings=Fittings)
    list_System_DuctData.append(list_dfData)


#######################################




### Initiate Empty DataFrame
df_DuctData = pd.DataFrame()

### Concatenate df[0] to df_DuctData
### df[0] is the DataFrame containing duct data from HTMLparse.py
for df in list_System_DuctData:
    df_DuctData = pd.concat([df_DuctData, df[0]])



##### BREAK - RUN TO HERE FOR HTML FILE TESTING #####


### Write data to excel file ###
with ExcelWriter(Input_excelpath, mode="a", if_sheet_exists='replace') as writer:
    df_DuctData.to_excel(writer, sheet_name=Input_sheetname_All)



### Write data to excel file ###
# with ExcelWriter(Input_excelpath, engine='openpyxl', mode="a", if_sheet_exists='replace', engine_kwargs={'keep_vba' : True}) as writer:
    # df_DuctData.to_excel(writer, sheet_name=Input_sheetname_All, )
    # df_DuctData_CP.to_excel(writer, sheet_name=Input_sheetname_CP)

# """