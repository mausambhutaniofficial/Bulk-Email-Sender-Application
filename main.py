from tkinter import *
from PIL import ImageTk
from tkinter import messagebox, filedialog
import os
import pandas as pd
import email_function       #another .py file in same directory
import time

class Bulk_email:
    def __init__(self,window):
        self.window=window
        self.window.title("Email Sender Application")
        self.window.geometry("1000x600+200+50")
        self.window.resizable(False,False)
        self.window.iconbitmap("images\email1.ico")
        self.window.config(bg="white")

        self.email_pic=ImageTk.PhotoImage(file=r"D:\source codes_Python\Bulk Email Sender App\images\email.png")
        self.settings_pic=ImageTk.PhotoImage(file=r"D:\source codes_Python\Bulk Email Sender App\images\settings.png")
        self.send_pic=ImageTk.PhotoImage(file=r"D:\source codes_Python\Bulk Email Sender App\images\send.png")
        self.browse_pic=ImageTk.PhotoImage(file=r"D:\source codes_Python\Bulk Email Sender App\images\browse.png")

        title=Label(self.window, text="Bulk Email Sender Panel",image=self.email_pic,compound=LEFT,padx=15,font=("Trebuchet MS",42,"bold","underline"), bg="#FCCAC0", anchor="w")
        title.place(x=0, y=0, relwidth=1)
        settings_btn=Button(self.window, image=self.settings_pic, cursor="hand2",activebackground="#FCCAC0",bg="#FCCAC0", bd=0 , command=self.settings_window)
        settings_btn.place(x=900, y=35)
        desc=Label(self.window, text="Use Excel File to Send the Bulk Emails at once with just one click. **Please Ensure the Email Column Name must be 'Email'",font=("Calibri (Body) ",12), bg="#F3A99A")
        desc.place(x=0, y=136, relwidth=1)


        self.var_choice=StringVar()
        single=Radiobutton(self.window, text="Single Mail",value="single",variable=self.var_choice,font=("times new roman",25),command=self.check_single_or_bulk,bg="white",activebackground="white")
        single.place(x=50, y=180)
        multiple=Radiobutton(self.window, text="Bulk Mail",value="multiple",variable=self.var_choice,font=("times new roman",25),command=self.check_single_or_bulk,bg="white",activebackground="white")
        multiple.place(x=250, y=180)
        self.var_choice.set("single")


        #-------------LABELS--------------------
        To=Label(self.window, text="To (Email Address):",font=("times new roman",15),bg="white")
        To.place(x=50, y=250)
        subj=Label(self.window, text="Subject:",font=("times new roman",15),bg="white")
        subj.place(x=50, y=300)
        msg=Label(self.window, text="Body:",font=("times new roman",15),bg="white")
        msg.place(x=50, y=350)

        #-------------ENTRY BOXES----------
        self.to_entry=Entry(self.window,width=45, font=("times new roman",15),bg="#EFEAE9")
        self.to_entry.place(x=230,y=250)

        self.subj_entry=Entry(self.window,width=60, font=("times new roman",15),bg="#EFEAE9")
        self.subj_entry.place(x=230,y=300)

        self.msg_entry=Text(self.window, font=("times new roman",15),bg="#EFEAE9")
        self.msg_entry.place(x=230,y=350,width=607,height=150)

        #----------STATUS LABELS------------
        self.total_mails=Label(self.window,font=("times new roman",15),bg="white")
        self.total_mails.place(x=50, y=550)
        
        self.sent_mails=Label(self.window,font=("times new roman",15),bg="white",fg="green")
        self.sent_mails.place(x=160, y=550)
        
        self.pending_mails=Label(self.window,font=("times new roman",15),bg="white",fg="orange")
        self.pending_mails.place(x=260, y=550)

        self.failed_mails=Label(self.window,font=("times new roman",15),bg="white",fg="red")
        self.failed_mails.place(x=400, y=550)


        #--------------BUTTONS------------------
        self.browse_btn=Button(self.window,command=self.browse_file_button, image=self.browse_pic,bg="white",activebackground="white",cursor="hand2",bd=0)
        self.browse_btn.config(state=DISABLED)
        self.browse_btn.place(x=690, y=215)

        clear_btn=Button(self.window, text="CLEAR",command=self.clear1,font=("times new roman",16,"underline"), activebackground="white",bg="white",bd=0,cursor="hand2",activeforeground="green", fg="#CB4F6D" )
        clear_btn.place(x=680, y=525)

        send_btn=Button(self.window, image=self.send_pic,command=self.send_email, activebackground="white",bg="white",cursor="hand2", bd=0)
        send_btn.place(x=770, y=500)
        self.check_file_exist()

    def browse_file_button(self):       #import filedialog from tkinter
        op=filedialog.askopenfile(initialdir='/',title='Select Excel File for Emails',filetypes=(("All Files","*.*"),("Excel Files",".xlsx")))
        print(op)
        if op!=None:
            data=pd.read_excel(op.name)
            #print(data['Email'])

            if 'Email' in data.columns:
                print("Email Field Exists")
                self.email_list=list(data['Email'])
                #print(email_list)
                #if there exist none values in email column:-
                new_list=[]
                for i in self.email_list:
                    if pd.isnull(i)==False:
                        #print(i)
                        new_list.append(i)
                self.email_list=new_list
                #print(self.email_list)           #LIST OF EMAILS WITHOUT NULL VALUES
                if len(self.email_list)>0:
                    self.to_entry.config(state=NORMAL)
                    self.to_entry.delete(0,END)
                    self.to_entry.insert(0, str(op.name.split("/")[-1]))
                    self.to_entry.config(state='readonly')
                    self.total_mails.config(text="Total: "+str(len(self.email_list)))
                    self.sent_mails.config(text="Sent: ")
                    self.pending_mails.config(text="Pending: ")
                    self.failed_mails.config(text="Failed: ")
                else:
                    messagebox.showerror("Error","Email Field does not Exists in this Excel file", parent=self.window)

            else:
                messagebox.showerror("Error","Email Field does not Exists in this Excel file", parent=self.window)








    def send_email(self):
        #print(self.to_entry.get(),self.subj_entry.get(),self.msg_entry.get("1.0",END))      #ncoz last wala Text box h
        x=len(self.msg_entry.get("1.0",END))            #by default its length is 1 when empty
        if self.to_entry.get()=="" or self.subj_entry.get()=="" or x==1:
            messagebox.showerror("Error","All fields are required", parent=self.window)
        else:
            if self.var_choice.get()=="single":
                status=email_function.email_send_function(self.to_entry.get(), self.subj_entry.get(), self.msg_entry.get("1.0", END), self.from_, self.pass_)
                if status=="s":
                    messagebox.showinfo("Success", "Email(s) Have Been Sent Successfully", parent=self.window)
                if status=="f":
                    messagebox.showerror("Error", "Email(s) Sending Failed, Try Again", parent=self.window)

            if self.var_choice.get()=="multiple":
                self.failed=[]
                self.s_counter=0
                self.f_counter=0
                for x in self.email_list:
                    status=email_function.email_send_function(x, self.subj_entry.get(), self.msg_entry.get("1.0", END), self.from_, self.pass_)
                    if status=="s":
                        self.s_counter+=1
                    if status=="f":
                        self.f_counter+=1
                    self.status_bar()
                    #time.sleep(1)
                messagebox.showinfo("Success", "Email(s) Have Been Sent Successfully", parent=self.window)

    
    def status_bar(self):
        self.total_mails.config(text="Status: "+str(len(self.email_list))+"=>>")
        self.sent_mails.config(text="Sent: "+str(self.s_counter))
        self.pending_mails.config(text="Pending: "+str(len(self.email_list) - (self.s_counter + self.f_counter)))
        self.failed_mails.config(text="Failed: "+str(self.f_counter))
        self.total_mails.update()
        self.sent_mails.update()
        self.pending_mails.update()
        self.failed_mails.update()


        #when Single is selected then Browse button should kept disabled
    def check_single_or_bulk(self):
        if self.var_choice.get()=="single":
            self.browse_btn.config(state=DISABLED)
            self.to_entry.config(state=NORMAL)
            self.to_entry.delete(0,END)
            self.clear1()
        else:
            self.var_choice.get()=="multiple"
            self.browse_btn.config(state=NORMAL)
            self.to_entry.delete(0,END)
            self.to_entry.config(state='readonly')



    def clear1(self):
        self.to_entry.config(state=NORMAL)
        self.to_entry.delete(0,END)
        self.subj_entry.delete(0,END)
        self.msg_entry.delete("1.0",END)
        self.var_choice.set("single")
        self.browse_btn.config(state=DISABLED)  
        #after pressing clear, all fields like total, failed etc, should be cleared also
        self.total_mails.config(text="")
        self.sent_mails.config(text="")
        self.pending_mails.config(text="")
        self.failed_mails.config(text="")


    def settings_window(self):
        self.check_file_exist()     #fn calling

        self.window2=Toplevel()         #it will create a child window

        self.window2.title("Settings")
        self.window2.geometry("700x350+350+90")
        self.window2.iconbitmap("images\email1.ico")
        self.window2.config(bg="white")
        self.window2.focus_force()      #focus on settings window when it is opened
        self.window2.grab_set()         #isse pehle settings window ko close krna hoga tab hi main window ko access kr skte ho

        title=Label(self.window2, text="Credential Settings",image=self.settings_pic,compound=LEFT,padx=15,pady=15,font=("Trebuchet MS",42,"bold","underline"), bg="#FCCAC0", anchor="w")
        title.place(x=0, y=0, relwidth=1)
        desc=Label(self.window2, text="Enter the Master Email adress and password from which to send the emails",font=("Calibri (Body) ",12), bg="#F3A99A")
        desc.place(x=0, y=106, relwidth=1)


        email=Label(self.window2, text="Email Address:",font=("times new roman",15),bg="white")
        email.place(x=50, y=150)
        passwrd=Label(self.window2, text="Password:",font=("times new roman",15),bg="white")
        passwrd.place(x=50, y=200)

        self.email_entry=Entry(self.window2,width=40, font=("times new roman",15),bg="#EFEAE9")
        self.email_entry.place(x=230,y=150)

        self.passwrd_entry=Entry(self.window2,width=20, font=("times new roman",15),bg="#EFEAE9", show="*")
        self.passwrd_entry.place(x=230,y=200)

        clear_btn=Button(self.window2,command=self.clear2, text="Clear",font=("times new roman",15,"underline"), activebackground="white",bg="white",bd=0,cursor="hand2",activeforeground="green", fg="#CB4F6D" )
        clear_btn.place(x=250, y=235)
        save_btn=Button(self.window2,command=self.save_setting, text="Save Credentials",font=("times new roman",15, "underline"), activebackground="white",bg="white",bd=0,cursor="hand2",activeforeground="red", fg="#CB4F6D" )
        save_btn.place(x=340, y=235)

        self.email_entry.insert(0,self.from_)
        self.passwrd_entry.insert(0,self.pass_)



    def check_file_exist(self):     #import os
        if os.path.exists("credentials.txt")==False:        #if file doesn't exists
            f=open('credentials.txt','w')
            f.write(self.email_entry.get()+" , "+ self.passwrd_entry.get())
            f.close()
        f2=open('credentials.txt','r')
        self.credential_list=[]
        for i in f2:
            #print(i)
            self.credential_list.append([i.split(",")[0],i.split(",")[1]])
        #print(self.credential_list)
        self.from_=self.credential_list[0][0]
        self.pass_=self.credential_list[0][1]
        #print(self.from_,self.pass_)

    def clear2(self):       #setting window fn
        self.email_entry.delete(0,END)
        self.passwrd_entry.delete(0,END)

    def save_setting(self):
        if self.email_entry.get()=="" or self.passwrd_entry.get()=="":
            messagebox.showerror("Error","All fields are required", parent=self.window2)
        else:
            f=open('credentials.txt','w')
            f.write(self.email_entry.get()+" , "+ self.passwrd_entry.get())
            f.close()
            messagebox.showinfo("Success","Credentials Saved Successfully", parent=self.window2)
            self.check_file_exist()


        self.window2.mainloop()

window=Tk()
obj=Bulk_email(window)
window.mainloop()

