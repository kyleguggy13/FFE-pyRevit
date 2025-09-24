# -*- coding: UTF-8 -*-
from datetime import datetime
from os import path, remove
from pyrevit import revit
from pyrevit.userconfig import user_config
from customOutput import def_syncLogPath

doc = __eventargs__.Document
filePath = doc.PathName

# getting central file name for log name
central_path = revit.query.get_central_path(doc)
try:
    # for rvt server
    lastBackslash_C = central_path.rindex("/")
except:
    # for other locations
    lastBackslash_C = central_path.rindex("\\")

# just the file name without the extension
central_file_name = central_path[lastBackslash_C:][:-4]

# getting local file name for tmp file name
try:
    lastBackslash_L = filePath.rindex("\\")
# for syncing detached central file
except:
    lastBackslash_L = filePath.rindex("/")
# just the file name without the extension
local_file_name = filePath[lastBackslash_L:][:-4]


fileExtension = central_path[-3:]

if fileExtension == "rvt":
    # tabulator between data to separte columns of the schedule
    separator = "\t" 
    try:
        # reading timestamp from tmp file
        try:
            # if parameter exists in config file
            try:
                syncLogPath = user_config.CustomToolsSettings.syncLogPath
            # if parameter doesnt exist in config file 
            except:
                syncLogPath = def_syncLogPath
            tmp_file_path = syncLogPath + "\\"+ local_file_name + "_Sync.tmp"
            # tmp_file_path = "L:\\customToolslogs\\syncTimeLogs\\"+ local_file_name + "_Sync.tmp"
        except:
            tmp_file_path = "\\\\Srv2\\Z\\customToolslogs\\syncTimeLogs\\"+ local_file_name + "_Sync.tmp"
        tmp_file = open(tmp_file_path, "r")
        start_time_string = tmp_file.read()
        # converting string to datetime
        start_time = datetime.strptime(start_time_string,"%Y-%m-%d %H:%M:%S")
        tmp_file.close()

        if path.exists(tmp_file_path):
            remove(tmp_file_path)

        # end time in seconds
        # round datetime to seconds (converting to string and then back to datetime object with correct)
        end_time_string_seconds = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        end_time = datetime.strptime(end_time_string_seconds,"%Y-%m-%d %H:%M:%S")


        timeDelta = end_time - start_time
        # print timeDelta

        user_name = doc.Application.Username

        # writing time to log file
        # if syncLogPath exists
        try:
            log_file = open(syncLogPath + "\\"+ central_file_name + "_Sync.log", "a")
            # log_file = open("L:\\customToolslogs\\syncTimeLogs\\"+ central_file_name + "_Sync.log", "a")
        # unc file path
        except:
            log_file = open("\\\\Srv2\\Z\\customToolslogs\\syncTimeLogs\\"+ central_file_name + "_Sync.log", "a")         
        log_file.write(end_time_string_seconds + separator + str(timeDelta) + separator + user_name + "\n")
        log_file.close()
    except:
         pass