
import pandas as pd
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
df = pd.read_csv("UAL_HV.csv")

users=df.User
departments = df.Access
    
    
def makeTitle(name):
    print("Harbourview PROD Access to: " + name + "\n")
    
    
def makeBody(name, access):
    print("Employee Name: " + name +
          "\n HarbourView Access\n" + 
          "Environment: PROD\n" +
          "Project(s): HarbourView Next Gen PROD\n" +
          "Privelages: NA\n"
          + "Permissions: " + access 
+ "\n----------------------------------------------------------")

def sendTicket():
    for i in range (len(users)):
        mail = outlook.CreateItem(0)
        mail.To = 'HARBOURVESTITEMAIL@harbourvest.com'
        mail.Subject = makeTitle(users[i])
        mail.Body = makeBody(users[i], departments[i])
        mail.Send()






