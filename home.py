from tkinter import *
from tkinter import messagebox
import os

home_tk=Tk()
home_tk.geometry("1920x1080")
home_tk.config(bg="#F5F5F5")
home_tk.title("Ajra Tex - Karur")
home_tk.configure(bg='white')

def login():
    en1 = c1.get()
    en2 = d1.get()
    
    if(en1 == "ajratex" and en2 == "password"):
        home_tk.destroy()
        os.system('python mainpage.py')
        
    
    else:
        messagebox.showerror(parent=home_tk, title="Invalid Credentials", message="Please enter valid username and password")
        
a = Label(home_tk,text="AJRA TEX - KARUR",font=("Arial", 30, "bold"),bg="white",fg="#1A374D").place(x=585,y=20)
b = Label(home_tk,text="LOGIN",font=("Arial",20,"bold"),bg="white",fg="#1A374D").place(x=720,y=130)
c = Label(home_tk,text="USERNAME",font=("Arial",20),bg="white",fg="#1A374D").place(x=700,y=250)
c1 = Entry(home_tk,font=("Calibri",18),bd=5,justify="center")
c1.place(x=650,y=300)
d = Label(home_tk,text="PASSWORD",font=("Arial",20),bg="white",fg="#1A374D").place(x=700,y=400)
d1 = Entry(home_tk,font=("Calibri",18),show="*",bd=5,justify="center")
d1.place(x=650,y=450)
bt1 = Button(home_tk,text="LOGIN",font=("Arial",15),bg="#1A374D",fg="#F5F5F5",width=10,height=1,command = login).place(x=720,y=550)

home_tk.mainloop()
