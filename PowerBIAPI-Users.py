#python -m pip install requests
import requests
#API call for all workspaces
AllWorkspace = "https://api.powerbi.com/v1.0/myorg/groups"
import pandas as pd
from datetime import datetime
import calendar
from tkinter import *
from tkinter import simpledialog
import webbrowser

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
        for w in resp_users_json:
            username.append(w["displayName"])
        for e in resp_users_json:
            email.append(e["identifier"])
        for r in resp_users_json:
            role.append(r["groupUserAccessRight"])
        
        reports_names = []
        reports = "https://api.powerbi.com/v1.0/myorg/groups/"+ workspace +"/reports"
        resp_reports = requests.get(reports, headers=headers)
        resp_reports_json = resp_reports.json()
        resp_reports_json = resp_reports_json["value"]
        for r in resp_reports_json:
            reports_names.append(r["name"])
        
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



#Define a callback function
def callback(url):
    webbrowser.open_new_tab(url)
gui = Tk()
pane = Frame(gui)
pane.pack(fill = BOTH, expand = True)

# Creatings a hpyer link Label Widget
link = Label(pane, text="https://learn.microsoft.com/en-us/rest/api/power-bi/groups/get-group-users",font=('ComicSans', 15), fg="blue", cursor="hand2")
link.pack()
link.bind("<Button-1>", lambda e:
callback("https://learn.microsoft.com/en-us/rest/api/power-bi/groups/get-group-users"))

L1 = Label(pane,text="Auth Key",font=('ComicSans', 15), fg="black")
L1.pack()
#Create an Entry Widget
entry= Entry(pane,width=40)
entry.pack()
#Create a button to display the text of entry widget
button= Button(pane, text="Enter", command=lambda: main(entry.get()))
button.pack(pady= 30)

gui.mainloop()