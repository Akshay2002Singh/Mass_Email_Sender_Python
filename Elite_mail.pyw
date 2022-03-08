from tkinter import filedialog
import threading
import os
from tkinter import *
from time import sleep
import pandas as pd

# 2003012005@IPEC.ORG.IN

# functions
def create_sheet(list,file_name):
    update_status(f"Creating Excel file")
    dataFrame = pd.DataFrame(list)
    dataFrame.to_excel(file_name,sheet_name='Sheet1',index = False,header=False)
def clear_url_box():
    From_email.set("")
    Password.set("")
    CSV_file.set("")
    Subject.set("")
    Body.set("")
def update_status(temp):
    statusvar.set(temp)
    sbar.update()
def download():
    # t1 = threading.Thread(target=createList, name='t1')
    # t1.start()
    pass

def change_path():
    folder = filedialog.askdirectory()
    os.chdir(folder)
    try:
        os.mkdir("Elite Email Generator")
    except:
        pass
    folder= os.path.join(folder,"Elite Email Generator")
    os.chdir(folder)

    download_loacation_display.update()

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
    CSV_file=StringVar()
    Subject=StringVar()
    Body=StringVar()
    CSV_file.set(f"Selected file :- {CSV_file.get()}")
    statusvar = StringVar()
    statusvar.set("Ready to download")
    # code to download a video
    heading1=Label(root,text="ELITE AKSHAY",font="calibre 20 bold",relief=RAISED,background="white",padx=10,pady=9)
    heading1.pack()
    space=Label(root,text="",font="calibre 2 bold")
    space.pack()
    heading2=Label(root,text="Elite Mass Email Sender",font="Times 25 bold",relief=RAISED,background="cyan",padx=10,pady=9,)
    heading2.pack()
    # input values 
    f1=Frame(root)
    f1.pack(side=TOP,fill=BOTH,expand=True,pady=10)
    name=Label(f1,text="Enter Your Email ID",font="calibre 18 bold italic",relief=FLAT,padx=8,pady=5).pack()
    user_name=Entry(f1,textvariable=From_email,font="calibre 15 normal",fg="blue",relief=SUNKEN,width=45)
    user_name.pack()
    Label(f1,text="Enter Your Password",font="calibre 18 bold italic",relief=FLAT,padx=8,pady=5).pack()
    Entry(f1,textvariable=Password,font="calibre 15 normal",fg="blue",relief=SUNKEN,width=35).pack()
    download_loacation_display=Label(f1,textvariable=CSV_file,font="calibre 10 bold italic",relief=FLAT,padx=18,pady=3)
    download_loacation_display.pack()
    space=Label(f1,text="",font="calibre 1 bold").pack()
    Label(f1,text="Subject Of Mail",font="calibre 18 bold italic",relief=FLAT,padx=8,pady=5).pack()
    Entry(f1,textvariable=Subject,font="calibre 15 normal",fg="blue",relief=SUNKEN,width=45).pack()
    Label(f1,text="Body",font="calibre 18 bold italic",relief=FLAT,padx=8,pady=5).pack()
    Entry(f1,textvariable=Body,font="calibre 15 normal",fg="blue",relief=SUNKEN,width=70).pack()

    
    download_btn=Button(f1,text="SELECT CSV",command=download,bd=5,fg="blue",font="calibre 18 bold")
    download_btn.pack(side = LEFT, expand = True, fill = X,pady=10)
    clear_url_btn=Button(f1,text="CLEAR URL",command=clear_url_box,bd=5,font="calibre 18 bold")
    clear_url_btn.pack(side = LEFT, expand = True, fill = X,pady=10)
    save_location=Button(f1,text="SEND",command=change_path,bd=5,fg="blue",font="calibre 18 bold")
    save_location.pack(side = LEFT, expand = True, fill = X,pady=10)
    # statusbar
    sbar = Label(root,textvariable=statusvar, relief=SUNKEN, anchor="w",padx=10,pady=7,background="cyan",fg="red",font="calibre 12 bold")
    sbar.pack(side=BOTTOM, fill=X)



    root.mainloop()