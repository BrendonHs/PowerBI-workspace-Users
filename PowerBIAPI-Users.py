#imports all of the package for the package check/download
import sys
import subprocess
import pkg_resources

required = {'requests', 'pandas','datetime'}
installed = {pkg.key for pkg in pkg_resources.working_set}
missing = required - installed

if missing:
    python = sys.executable
    subprocess.check_call([python, '-m', 'pip', 'install', *missing], stdout=subprocess.DEVNULL)

#imports all of the packages required to run everything
import requests
import pandas as pd
from datetime import datetime
import calendar
from tkinter import *
import tkinter as tk
import webbrowser

#API call for all workspaces
AllWorkspace = "https://api.powerbi.com/v1.0/myorg/groups"
def main(authToken):
    headers = {
            'Authorization': 'Bearer ' + authToken,
            'Content-Type': 'application/json'
        }

    #calls the API 
    resp = requests.get(AllWorkspace, headers=headers)
    resp_json = resp.json()
    resp_json = resp_json["value"]

    workspaceid = []
    workspacename = []
    #Appends all of the workspace ID to a list
    for w in resp_json:
        workspaceid.append(w['id'])
        workspacename.append(w['name'])

    Master_df = pd.DataFrame()

    workspace_num = 0
    #Appends all of the users for an individual workspace to a list 
    for workspace in workspaceid:
        username = []
        email = []
        role =[]
        user = "https://api.powerbi.com/v1.0/myorg/groups/"+ workspace +"/users"
        resp_users = requests.get(user, headers=headers)
        resp_users_json = resp_users.json()
        resp_users_json = resp_users_json["value"]
        #appends the jason file into a list
        for w in resp_users_json:
            username.append(w["displayName"])
        for e in resp_users_json:
            email.append(e["identifier"])
        for r in resp_users_json:
            role.append(r["groupUserAccessRight"])
        
        #Puts all of the reports into a list 
        reports_names = []
        reports = "https://api.powerbi.com/v1.0/myorg/groups/"+ workspace +"/reports"
        resp_reports = requests.get(reports, headers=headers)
        resp_reports_json = resp_reports.json()
        resp_reports_json = resp_reports_json["value"]
        for r in resp_reports_json:
            reports_names.append(r["name"])
        
        #creates DF for for export into an CSV 
        d=[{'workspace':workspacename[workspace_num]}]
        df = pd.DataFrame(d)
        reports_df = pd.DataFrame({'reports_names':reports_names})
        df = pd.concat([df, reports_df], axis=1) 
        names_df = pd.DataFrame({"displayName" : username})
        df = pd.concat([df, names_df], axis=1) 
        email_df = pd.DataFrame({"email" : email})
        df = pd.concat([df, email_df], axis=1) 
        role_df = pd.DataFrame({"role" : role})
        df = pd.concat([df, role_df], axis=1) 
        Master_df = pd.concat([Master_df,df])
        workspace_num += 1

    currentMonth = datetime.now().month
    Master_df.to_csv(str(calendar.month_name[currentMonth])+"PowerBIAudit.csv")

#The instructions for how this things works
instruction= [
            "A. Create a Microsoft learn account.",
            "B. Once a Microsoft learn account has been created click on the green try it button for the Microsoft learn page and sign in again",
            "C. Confirm your account in the new side window",
            "D. Go to the section that has the header Request Review and copy the Authorization number.",
            "E. Do no copy the word “Bearer” only the numbers/letters behind it",
            "F. Copy and paste the authorization key into it",
            "H. A csv file will pop in the same folder as where the script is located with the information",
            "I. Close the Program by Clicking Program button"]

#Define a callback function to open up a webpage
def callback(url):
    webbrowser.open_new_tab(url)
def close():
   gui.quit()

gui = Tk()
gui.title("Power BI API Users")

# Creatings a hpyer link Label Widget
link = Label(gui, text="Link: https://learn.microsoft.com/en-us/rest/api/power-bi/groups/get-group-users",font=(15), fg="blue", cursor="hand2")
link.pack()
link.bind("<Button-1>", lambda e:
callback("https://learn.microsoft.com/en-us/rest/api/power-bi/groups/get-group-users"))

#Instructions Text Box
title = Label(gui, text="Instructions",font='Comic_Sans_MS 15 bold', fg="black",)
title.pack()
count = 1
listbox = Listbox(gui, width = 120,height = 8,font=(12), fg="black")
for i in instruction:
    listbox.insert(count,i)
    count += 1
listbox.pack()

#Label for Auth Key Text
L1 = Label(gui,text="Auth Key",font=(15), fg="black")
L1.pack()
#Create an Entry Widget
entry= Entry(gui,width = 40,font=(8))
entry.pack()
#Create a button to display the text of entry widget
button= Button(gui, text="Enter",font=(8), command=lambda:[main(entry.get()),close])
button.pack(pady = 10)
buttonclose = Button(gui, text="Close Program",font=(8), command=close)
buttonclose.pack(pady = 10)

gui.mainloop()