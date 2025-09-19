# -*- coding: utf-8 -*- 
from pyrevit import coreutils
from pyrevit import output

# colors for chart.js graphs
colors = 10*["#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb",
            "#4d4d4d","#000000","#fff0f2","#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#e97800","#a6c844",
            "#4d4d4d","#fff0d9","#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#4d4d4d","#e97800","#a6c844",
            "#fff0e6","#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#4d4d4d","#fff0e6","#e97800","#a6c844",
            "#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#4d4d4d","#9988bb","#4d4d4d","#e97800","#a6c844",
            "#4d4d4d","#fff0d9","#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#4d4d4d","#e97800","#a6c844",
            "#4d4d4d","#fff0d9","#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#4d4d4d","#e97800","#a6c844",
            "#4d4d4d","#fff0d9","#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#4d4d4d","#e97800","#a6c844",
            "#4d4d4d","#fff0d9","#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#4d4d4d","#e97800","#a6c844",
            "#4d4d4d","#fff0d9","#ffc299","#ff751a","#cc5200","#ff6666","#ffd480","#b33c00","#ff884d","#d9d9d9","#9988bb","#4d4d4d","#e97800","#a6c844",]

# list of Warnings rated as critical
criticalWarnings = ['Elements have duplicate "Type Mark" values',
    'There are identical instances in the same place',
    'Room Tag is outside of its Room',
    'Multiple Rooms are in the same enclosed region',
    'Multiple Areas are in the same enclosed region',
    'One element is completely inside another',
    'Room is not in a properly enclosed region',
    'Room separation line is slightly off axis and may cause inaccuracies',
    'Area is not in a properly enclosed region',
    "Rectangular opening doesn't cut its host",
    'Elements have duplicate "Number" values',]

# default paths for settings
# this can be overriden by user in customToolsSettings.py
# or by company config file
def_hookLogs = "L:\\customToolslogs\\hooksLogs"
def_revitBuildLogs = "L:\\customToolslogs\\versions.log"
# 2021.1.6, 2022.1.3, 2023.1.1 - this can be overriden in customToolsSettings
def_revitBuilds = "20220123_1515(x64), 20220520_1515(x64), 20221122_1550(x64)"
def_massMessagePath = "L:\\_i\\CTmassMessage\\mass_message.html"
def_syncLogPath = "L:\\customToolslogs\\syncTimeLogs"
def_openingLogPath = "L:\\customToolslogs\\openingTimeLogs"
def_dashboardsPath = "L:\\powerBI"
def_language = 1
def_doorUnflipped = "L"
def_doorFlipped = "R"
def_windowUnflipped = "1" 
def_windowFlipped = "0"
def_wiki = ""
def_standardWorksets = "Shared Levels & Grids, Workset1, Hidden, 3DtextRoomData"

# read the company config file if it does exist
def company_conf():
    # constructing the drive list
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    drive_list = []
    for i in alphabet:
        drive_name = "".join((i,":\\"))
        drive_list.append(drive_name)

    config_values = {}
    # checking every single letter as drive location A:\ B:\ C:\ etc.
    for drive in drive_list:
        try:
            company_config_file = open(drive + "ct_config.ini","r")
            configSetting = (company_config_file.readlines())
            for line in configSetting:
                # separating values from lines
                keys = line.split(" = ")
                # remove new lines - enter
                keys[1] = keys[1].split("\n")[0]
                # add values to dictionary
                config_values.update({keys[0]:keys[1]})
            company_config_file.close()
        # if there is no company config file, skip this step and use default values
        except:
            pass
    return config_values

# formating time in seconts to HHMMSS format
def hmsTimer(timerSeconds):
    # for treating formating of pyrevit timer function
    from math import floor
    seconds = round(timerSeconds,2)
    if seconds<60:
        hms = str(seconds)+" seconds"
    elif seconds<3600:
        minutes = int(floor(seconds//60))
        seconds = seconds%60
        hms = str(minutes)+" min "+str(seconds)+" seconds"
    else:
        hours = seconds//3600
        minutes = int((seconds%3600)//60)
        seconds = seconds%60
        if minutes ==0:
            hms = str(hours)+" h "+str(seconds)+" seconds"
        else:
            hms = str(hours)+" h "+str(minutes)+" min "+str(seconds)+" seconds"
    claim = "Transaction took "+hms
    return claim

# gets name of the current document
def file_name_getter(doc):
    file_path = doc.PathName
    # trying all cases, for worshared, not worshared and detached files
    try:
        file_name = file_path[file_path.rindex("/"):]
    except:
        try:
            file_name = file_path[file_path.rindex("\\")+1:]
        except:
            file_name = file_path
    return(file_name)

# setting icon for output window
def ct_icon(output):
    import os
    appdataPath = os.getenv('APPDATA')
    iconPath = appdataPath + '\\pyRevit\\Extensions\\CustomTools.extension\\CustomToolsLogo.PNG'
    output.set_icon(iconPath)

# creating mass message url
def mass_message_url(output):
    from customOutput import def_massMessagePath
    from pyrevit.userconfig import user_config
    from os import path
    import os
    appdataPath = os.getenv('APPDATA')
    # server version of massmessage
    # if parameter exists in config file
    try:
        url = user_config.CustomToolsSettings.massMessagePath
    # if parameter doesnt exist in config file
    except:
        url = def_massMessagePath

    # url = "L:\\_i\\CTmassMessage\\mass_message.html"
    # url_unc = "\\\\Srv\\Z\\_i\\CTmassMessage\\mass_message.html"
    if path.exists(url):
        # output.open_url(url)
        return url
    # elif path.exists(url_unc):
    #     return url_unc
    # offline hardcoded version of massmessage
    else:
        # offline content of mass message
        return appdataPath + '\\pyRevit\\Extensions\\CustomTools.extension\\mass_message\\mass_message.html'

# highlights text using html string with css
def text_highligter(a):
    content = str(a)
    html_code = "<p class='elementlink'>"+content+"</p>"
    return coreutils.prepare_html_str(html_code)

# makes mailto link in output window
def mailto(a):
    content = str(a)
    html_code = '<a href=mailto:"'+ content +'" target="_blank" style="text-decoration: none; color: black; font-weight: bold;">'+ content +'</a>'
    # html_code = '<a href=mailto:"'+ content +'" target="_blank">'+ content +'</a>'
    return coreutils.prepare_html_str(html_code)

# makes html link tag
def linkMaker(a,title):
    content = str(a)
    html_code = '<a href="'+content+'">'+ title +'</a>'
    return coreutils.prepare_html_str(html_code)

# views image in output window
def imageViewer(html_code):
    # sample_code = "<img src='https://i.ytimg.com/vi/SfLV8hD7zX4/maxresdefault.jpg' width=50%>"
    print(coreutils.prepare_html_str(html_code))

# heading in html
def heading(text, size):
    # sample_code = "<h2>Heading</h2>"
    size_str = str(size)
    html_code = "<h" + size_str +">" + text + "</h" + size_str +">"
    print(coreutils.prepare_html_str(html_code))