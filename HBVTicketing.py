import pandas as pd
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

# reads user excel spreadsheet, must have column 'User" which is full name of person,
# and 'Access' column, which is a column describing access level, e.g. 'HVW Marketing BASIC' 

print("Input CSV must have a 'User' column which is the full name of a person needing access, e.g. John Doe \n" +
          "and an 'Access' column, which is a column describing access levels, like 'HVW Marketing BASIC'")

df = pd.read_csv(input("Enter the exact filepath of the csv speadsheet with users and their access levels: \n" + 
                       "e.g. C:/Users/johndoe/Documents/UAL.csv \n"))
# df = pd.read_csv('C:/Users/kmoran/Documents/UAL_HV.csv')

# establishes dataframes of the pertinent columns
users = df.User
departments = df.Access
    
# defines the subject of the ticket 
def makeSubject(name):
    return("HarbourView PROD Access Request For: " + name + "\n")
    
# defines the message body of the ticket     
def makeBody(name, access):
    return("Employee Name: " + name +
          "\n\n HarbourView Access\n\n" + 
          "Environment: PROD\n\n" +
          "Project(s): HarbourView Next Gen PROD\n\n" +
          "Privelages: NA\n\n"
          + "Permissions: " + access)

# sends the ticket via email to IT utilizing the outlook desktop client
def sendTickets():
    for i in range(len(users)):
        subject = makeSubject(users[i])
        body = makeBody(users[i], departments[i])
    
        mail = outlook.CreateItem(0)
        mail.To = 'helpdesk@harbourvest.com'        
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        print("Access Request Sent for: " + users[i])
    
        
def main():
    sendTickets()
    
if __name__ == '__main__': 
    main()
