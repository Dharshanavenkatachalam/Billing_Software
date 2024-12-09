from tkinter import *
import os

mp_tk=Tk()
mp_tk.geometry("1920x1080")
mp_tk.title("Home Page - Ajra Tex")
mp_tk.configure(bg='white')

def invoice():
    os.system('python invoice.py')
    
def delivery_slip():
    os.system('python delivery_slip.py')

def expence_tracker():
    os.system('python expence_tracker.py')

def report_analysis():
    os.system('python report_analysis.py')
    
a = Label(mp_tk,text="AJRA TEX - KARUR",font=("Arial", 30, "bold"),bg="white",fg="#1A374D").place(x=585,y=20)
bt1 = Button(mp_tk,text="INVOICE",font=("Arial",15),bg="#1A374D",fg="#F5F5F5",height=5,width=20,command=invoice).place(x=400,y=250)
bt2 = Button(mp_tk,text="DELIVERY SLIP",font=("Arial",15),bg="#1A374D",fg="#F5F5F5",height=5,width=20,command=delivery_slip).place(x=900,y=250)
bt3 = Button(mp_tk,text="EXPENSE TRACKER",font=("Arial",15),bg="#1A374D",fg="#F5F5F5",height=5,width=20,command=expence_tracker).place(x=400,y=500)
bt4 = Button(mp_tk,text="REPORT & ANALYSIS",font=("Arial",15),bg="#1A374D",fg="#F5F5F5",height=5,width=20,command=report_analysis).place(x=900,y=500)

mp_tk.mainloop()
