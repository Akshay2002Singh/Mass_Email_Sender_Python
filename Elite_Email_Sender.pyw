import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import threading
from tkinter.filedialog import askopenfilename
from tkinter import *
from time import sleep
from turtle import left
import pandas as pd
import os
import webbrowser


# functions
def create_sheet():
    update_status("Creating Email list from Excel file")
    file = CSV_file.get()
    try:
        df = pd.read_excel (file,sheet_name='Sheet1',header=None)
        df = df.iloc[:,0]
        df = df.dropna(axis=0)
        list_to = df[:].tolist()
        update_status("Senders list ready")
        return list_to
    except:
        update_status("Some error with sheet")
        sleep(1)
        update_status("Ready To Go On a Ride")
        return []

    
    
def clear_url_box():
    From_email.set("")
    Password.set("")
    CSV_file.set("")
    Subject.set("")
    textarea.delete(1.0,END)
    CSV_file.set("")
    CSV_file_loaction.set(f"Selected file :- {CSV_file.get()}")
    attachment_file.set("")
    attachment_file_loaction.set(f"Selected file :- {attachment_file.get()}")

def update_status(temp):
    statusvar.set(temp)
    sbar.update()

def send_start():
    t1 = threading.Thread(target=Start_task, name='t1')
    t1.start()
    
def Start_task():
    update_status("On a Ride")
    sleep(1)
    send["state"] = "disabled"
    send["text"] = "WAIT"
    Select_file["state"] = "disabled"
    Select_file["text"] = "WAIT"
    temp_file = attachment_file.get()
    is_attachment = 0

    if temp_file == "":
        is_attachment = 0
    else:
        is_attachment = 1
    email_from = From_email.get()
    email_password = Password.get()
    email_subject = Subject.get()
    email_body = textarea.get(1.0,END)

    # print(f"{email_from}\n{email_password}\n{email_subject}\n{email_body}")
    if email_from == "" or email_password == "" or email_body == "" or email_subject == "" :
        update_status("Values can not be empty")
        sleep(1)
        update_status("Ready To Go On a Ride")
    else:
        to_list = create_sheet()
        # print(to_list)
        size_total = len(to_list)
        if size_total > 0:
            try:
                update_status("Getting Ready to Send msg")
                # instance of MIMEMultipart
                msg = MIMEMultipart()

                # storing the senders email address
                msg['From'] = email_from

                # storing the subject
                msg['Subject'] = email_subject

                # string to store the body of the mail
                body = email_body

                # attach the body with the msg instance
                msg.attach(MIMEText(body, 'plain'))

                try:
                    if is_attachment == 1:
                        # open the file to be sent
                        update_status("Getting Ready to attach file")
                        
                        head, filename = os.path.split(temp_file)
                        attachment = open(temp_file, "rb")

                        # instance of MIMEBase and named as p
                        p = MIMEBase('application', 'octet-stream')

                        # To change the payload into encoded form
                        p.set_payload((attachment).read())

                        # encode into base64
                        encoders.encode_base64(p)

                        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

                        # attach the instance 'p' to instance 'msg'
                        msg.attach(p)
                except:
                    update_status("Error in attachment")
                    sleep(1)
                    update_status("continue ...")


                # creates SMTP session
                s = smtplib.SMTP('smtp.gmail.com', 587)

                # start TLS for security
                s.starttls()

                # Authentication
                s.login(email_from, email_password)

                # Converts the Multipart msg into a string
                text = msg.as_string()
                update_status("Sending msg")
                i=1
                for sender in to_list:
                    try:
                        update_status(f"Sending {i}/{size_total}")
                        i = i + 1
                        # sending the mail
                        msg['To'] = sender
                        s.sendmail(email_from, sender, text)
                    except:
                        update_status("1 Failed")
                sleep(2)

                # terminating the session
                s.quit()
            except:
                update_status("Some Error In Login")
                sleep(1)
                update_status("Ready To Go On a Ride")
                # clear_url_box()
        else:
            update_status("Some Error With Sender List")
            sleep(1)
            update_status("Ready To Go On a Ride")
            clear_url_box()
            send["state"] = "active"
            send["text"] = "SEND"
            Select_file["state"] = "active"
            Select_file["text"] = "SELECT CSV"
            return
        
    update_status("Ready To Go On a Ride")
    send["state"] = "active"
    send["text"] = "SEND"
    Select_file["state"] = "active"
    Select_file["text"] = "SELECT CSV"
    clear_url_box()

def Select_file():
    file=askopenfilename(defaultextension=".xlsx",filetypes =[('Excel Files', '*.xlsx'),('CSV Files', '*.csv')])
    CSV_file.set(file)
    CSV_file_loaction.set(f"Selected Senders file :- {CSV_file.get()}")

def attach_file():
    file_attachment=askopenfilename(defaultextension=".txt",filetypes=[("All Files","*")])
    attachment_file.set(file_attachment)
    attachment_file_loaction.set(f"Selected attachment file :- {attachment_file.get()}")

# function to open browser
def callback(url):
   webbrowser.open_new_tab(url)


# main body
if __name__=="__main__":
    root = Tk()
    # window size
    root.title("Elite Mass Email Sender")
    root.geometry("950x720")
    root.minsize(890,720)
    

    # Variables
    From_email = StringVar()
    Password=StringVar()
    CSV_file_loaction=StringVar()
    CSV_file=StringVar()
    attachment_file_loaction=StringVar()
    attachment_file=StringVar()
    Subject=StringVar()
    # Body=StringVar()

    statusvar = StringVar()
    statusvar.set("Ready To Go On a Ride")
    CSV_file_loaction.set(f"Selected Senders file :- {CSV_file.get()}")
    attachment_file_loaction.set(f"Selected attachment file :- {attachment_file.get()}")

    
    heading1=Label(root,text="ELITE AKSHAY",font="calibre 20 bold",relief=RAISED,background="white",padx=10,pady=9)
    heading1.pack()
    space=Label(root,text="",font="calibre 2 bold")
    space.pack()
    heading2=Label(root,text="Elite Mass Email Sender",font="Times 25 bold",relief=RAISED,background="cyan",padx=10,pady=9,)
    heading2.pack()
    
    
    f1=Frame(root)
    f1.pack(side=TOP,fill=BOTH,expand=True,pady=10)
    name=Label(f1,text="Enter Your Email ID",font="calibre 18 bold italic",relief=FLAT,padx=8,pady=5).pack()
    user_name=Entry(f1,textvariable=From_email,font="calibre 15 normal",fg="blue",relief=SUNKEN,width=45)
    user_name.pack()
    Label(f1,text="Enter App Password",font="calibre 18 bold italic",relief=FLAT,padx=8,pady=5).pack()
    Entry(f1,textvariable=Password,font="calibre 15 normal",fg="blue",relief=SUNKEN,width=35).pack()
    #Create a Label to display the link to generate app password
    api_link = Label(f1, text="Don't have App Password? Click here",font=('Helveticabold', 15), fg="blue", cursor="hand2")
    api_link.pack()
    api_link.bind("<Button-1>", lambda e: callback("https://support.google.com/mail/answer/185833?hl=en"))
    senders_file_location=Label(f1,textvariable=CSV_file_loaction,font="calibre 10 bold italic",relief=FLAT,padx=18,pady=3)
    senders_file_location.pack()
    attachment=Label(f1,textvariable=attachment_file_loaction,font="calibre 10 bold italic",relief=FLAT,padx=18,pady=3)
    attachment.pack()
    space=Label(f1,text="",font="calibre 1 bold").pack()
    Label(f1,text="Subject Of Mail",font="calibre 18 bold italic",relief=FLAT,padx=8,pady=5).pack()
    Entry(f1,textvariable=Subject,font="calibre 15 normal",fg="blue",relief=SUNKEN,width=45).pack()
    Label(f1,text="Body",font="calibre 18 bold italic",relief=FLAT,padx=8,pady=5).pack()
    # body
    body_frame = Frame(f1)
    body_frame.pack()
    textarea=Text(body_frame,font="Ariel 13 normal",height=6)
    textarea.pack(side='left',expand=False)
    Scroll =Scrollbar(body_frame)
    Scroll.pack(side=RIGHT,fill=Y)
    Scroll.config(command=textarea.yview)
    textarea.config(yscrollcommand=Scroll.set)
    
    Select_file=Button(f1,text="SELECT CSV",command=Select_file,bd=5,fg="blue",font="calibre 18 bold")
    Select_file.pack(side = LEFT, expand = True, fill = X,pady=3)
    clear_url_btn=Button(f1,text="CLEAR URL",command=clear_url_box,bd=5,font="calibre 18 bold")
    clear_url_btn.pack(side = LEFT, expand = True, fill = X,pady=3)
    send=Button(f1,text="SEND",command=send_start,bd=5,fg="blue",font="calibre 18 bold")
    send.pack(side = LEFT, expand = True, fill = X,pady=3)
    attachment_btn=Button(root,text="ATTACH FILE",command=attach_file,bd=5,fg="blue",font="calibre 18 bold")
    attachment_btn.pack(side=TOP)

    # statusbar 
    sbar = Label(root,textvariable=statusvar, relief=SUNKEN, anchor="w",padx=10,pady=7,background="cyan",fg="red",font="calibre 12 bold")
    sbar.pack(side=BOTTOM, fill=X)



    root.mainloop()