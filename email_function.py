#importing this file into main file

import smtplib
def email_send_function(to_,subj_,msg_,from_,pass_):
    #print(to_,subj_,msg_,from_,pass_)
    #CREATING SESSION FOR GMAIL
    s=smtplib.SMTP("smtp.gmail.com",587)
    s.starttls()    #it works the transport layer, provides security of the msg
    s.login(from_, pass_)
    msg=f"Subject: {subj_}\n\n{msg_}"       #ADDING subject with body in our gmail msg
    s.sendmail(from_, to_, msg)
    x=s.ehlo()    #returns status of te mail        if s.ehlo() =250 SUCCESS oterwise FAIL
    if x[0]==250:
        return "s"
    else:
        return "f"
    
    s.close()       #session closed
