
import pandas as pd
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

for i in range (len(users)):
    makeTitle(users[i])
    makeBody(users[i], departments[i])






