import pandas as pd
data=pd.read_excel("with_email.xlsx")
#print(data['Email'])

if 'Email' in data.columns:
    print("Email Field Exists")
    email_list=list(data['Email'])
    print(email_list)
#if there exist none values in email column:-
    new_list=[]
    for i in email_list:
        if pd.isnull(i)==False:
            #print(i)
            new_list.append(i)
    email_list=new_list
    print(email_list)           #LIST OF EMAILS WITHOUT NULL VALUES

else:
    print("Email Field does not Exists")
    
