
import pandas as pd
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

# reads user excel spreadsheet, must have column 'User" which is full name of person,
# and 'Access' column, which is a column describing access level, e.g. 'HVW Marketing BASIC' 
df = pd.read_csv("UAL_HV.csv")

# establishes dataframes of the pertinent columns
users = df.User
departments = df.Access
    
# defines the subject of the ticket 
def makeSubject(name):
    print("Harbourview PROD Access to: " + name + "\n")
    
# defines the message body of the ticket     
def makeBody(name, access):
    print("Employee Name: " + name +
          "\n HarbourView Access\n" + 
          "Environment: PROD\n" +
          "Project(s): HarbourView Next Gen PROD\n" +
          "Privelages: NA\n"
          + "Permissions: " + access)

# sends the ticket via email to IT utilizing the outlook desktop client
def sendTickets():
    for i in range (len(users)):
        mail = outlook.CreateItem(0)
        mail.To = 'HARBOURVESTITEMAIL@harbourvest.com'
        mail.Subject = makeSubject(users[i])
        mail.Body = makeBody(users[i], departments[i])
        mail.Send()
        
def main():
    sendTickets()
    
if __name__ == '__main__': 
    main()






