import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

# reads user excel spreadsheet, must have column 'User" which is full name of person,
# and 'Access' column, which is a column describing access level, e.g. 'HVW Marketing BASIC' 


name = (input("Enter the first and last name of the user needing access: \n"))
access = (input("Enter the access permissions for this user: (e.g. HVW Marketing BASIC)\n"))

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
def sendTicket():
    subject = makeSubject(name)
    body = makeBody(name, access)
    
    mail = outlook.CreateItem(0)
    mail.To = 'helpdesk@harbourvest.com'        
    mail.Subject = subject
    mail.Body = body
    mail.Send()
    print("Access Request Ticket Sent for: " + name)
    
        
def main():
    sendTicket()
    
if __name__ == '__main__': 
    main()
