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
    User.set("")
    NoOfStudent.set(0)
def update_status(temp):
    statusvar.set(temp)
    sbar.update()
def download():
    t1 = threading.Thread(target=createList, name='t1')
    t1.start()
def createList():
    download_btn.config(text="Wait",command=None)
    username = User.get()
    no_of_student= NoOfStudent.get()
    update_status("Creating File")
    list = []
    if check(username) == 0:
        update_status("Enter Valid Email")
        sleep(0.2)
        download_btn.config(text="CREATE CSV",command=download)
        update_status("Ready to Download")
        return
    elif no_of_student < 1:
        update_status("Enter Valid Number Of Students")
        sleep(0.2)
        download_btn.config(text="CREATE CSV",command=download)
        update_status("Ready to Download")
        return
    pin = username.split("@")
    pin = int(pin[0])
    pin = pin // 1000
    pin = pin * 1000
    t_noOfStudent=0
    for i in range(no_of_student+50):
        temp_mail=f"{pin+i}@ipec.org.in"
        if check(temp_mail):
            t_noOfStudent= t_noOfStudent+1
            list.append(temp_mail)
        update_status(f"Creating File ({t_noOfStudent} out of {no_of_student})")
        if t_noOfStudent == no_of_student:
            break
        
    create_sheet(list,f"{username}_{t_noOfStudent}.xlsx")
    download_btn.config(text="CREATE CSV",command=download)
    update_status("Ready to Download")

def change_path():
    folder = filedialog.askdirectory()
    os.chdir(folder)
    try:
        os.mkdir("Elite Email Generator")
    except:
        pass
    folder= os.path.join(folder,"Elite Email Generator")
    os.chdir(folder)
    download_path.set(folder)
    downloads_location.set(f"Save path :- {folder}")
    download_loacation_display.update()

# main body
if __name__=="__main__":
    root = Tk()
    # window size
    root.title("Elite Insta Profile Photo Downloader")
    root.geometry("850x520")
    root.minsize(800,400)
    

    # Variables
    User = StringVar()
    NoOfStudent=IntVar()
    downloads_location=StringVar()
    download_path=StringVar()
    download_path.set(os.getcwd())
    downloads_location.set(f"Save path :- {download_path.get()}")
    statusvar = StringVar()
    statusvar.set("Ready to download")
    # code to download a video
    heading1=Label(root,text="ELITE AKSHAY",font="calibre 40 bold",relief=RAISED,background="white",padx=10,pady=9)
    heading1.pack()
    space=Label(root,text="",font="calibre 2 bold")
    space.pack()
    heading2=Label(root,text="AUTO EMAIL GENERATOR",font="Times 25 bold",relief=RAISED,background="cyan",padx=10,pady=9,)
    heading2.pack()
    f1=Frame(root)
    f1.pack(side=TOP,fill=BOTH,expand=True,pady=10)
    name=Label(f1,text="ENTER COLLEGE MAIL ID OF ANY ONE STUDENT",font="calibre 20 bold italic",relief=FLAT,padx=8,pady=5,)
    name.pack()
    space=Label(f1,text="",font="calibre 1 bold")
    space.pack()
    user_name=Entry(f1,textvariable=User,font="calibre 25 normal",fg="blue",relief=SUNKEN)
    user_name.pack()
    Label(f1,text="ENTER TOTAL NUMBER OF STUDENT IN THAT BRANCH",font="calibre 16 bold italic",relief=FLAT,padx=8,pady=5,).pack()
    Entry(f1,textvariable=NoOfStudent,font="calibre 20 normal",fg="blue",relief=SUNKEN).pack()
    download_loacation_display=Label(f1,textvariable=downloads_location,font="calibre 10 bold italic",relief=FLAT,padx=18,pady=3)
    download_loacation_display.pack()
    
    download_btn=Button(f1,text="CREATE CSV",command=download,bd=5,fg="blue",font="calibre 18 bold")
    download_btn.pack(side = LEFT, expand = True, fill = X,pady=10)
    clear_url_btn=Button(f1,text="CLEAR URL",command=clear_url_box,bd=5,font="calibre 18 bold")
    clear_url_btn.pack(side = LEFT, expand = True, fill = X,pady=10)
    save_location=Button(f1,text="SAVE LOCATION",command=change_path,bd=5,fg="blue",font="calibre 18 bold")
    save_location.pack(side = LEFT, expand = True, fill = X,pady=10)
    # statusbar
    sbar = Label(root,textvariable=statusvar, relief=SUNKEN, anchor="w",padx=10,pady=7,background="cyan",fg="red",font="calibre 12 bold")
    sbar.pack(side=BOTTOM, fill=X)



    root.mainloop()