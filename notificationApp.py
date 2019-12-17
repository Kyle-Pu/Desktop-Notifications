#!/usr/bin/python
import gi
gi.require_version('Notify', '0.7')
from gi.repository import Notify

# Manage Time!
from datetime import datetime
import time

# Spreadsheet Management!
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

Notify.init("Hi")

data = pd.read_excel("Notifications.xlsx")

timeList = []
titleList = []
messageList = []

for note in data.index:
    timeList.append(data['Time'][note].strftime("%H:%M:%S"))
    titleList.append(data['Title'][note])
    messageList.append(data['Message'][note])
noteShown = False # Keep track of if notification was shown yet

print("Notifications will come! Just leave me in the background :D")

while int(noteShown) == 0:
    for ind in range(len(timeList) - 1):
        if datetime.now().strftime("%H:%M:%S") == timeList[ind]:
            Notify.Notification.new(timeList[ind], titleList[ind], "dialog-information").show()
            time.sleep(1) # Prevent notification repeats
