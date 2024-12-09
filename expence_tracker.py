from tkinter import *
from datetime import date
import mysql.connector as sql
from docxtpl import DocxTemplate
from tkinter import messagebox

et_tk=Tk()
et_tk.geometry("1920x1080")
et_tk.config(bg="#F5F5F5")
et_tk.title("Expence Tracker - Ajra Tex")
et_tk.configure(bg='white')

mycon=sql.connect(host="localhost",user="root",passwd="gokul123",database="ajra")
cursor=mycon.cursor()

def display_block():
    today = date.today()
    show = ("SELECT amount FROM accounts where date = '{date}'".format(date=today))
    cursor.execute(show)
    today_data=cursor.fetchall()

    today_balance_amount = 0.0
    today_from_account = 0.0
    today_to_account = 0.0
    for i in today_data:
        if(i[0] < 0):
            today_from_account += i[0]
        else:
            today_to_account += i[0]

    today_balance_amount = today_to_account + today_from_account
    if(today_from_account < 0):
        today_from_account *= (-1)
        
    month = date.today().month
    show = ("SELECT amount FROM accounts where month(date) = '{month}'".format(month=month))
    cursor.execute(show)
    month_data=cursor.fetchall()
    
    month_balance_amount = 0.0
    month_from_account = 0.0
    month_to_account = 0.0
    for i in month_data:
        if(i[0] < 0):
            month_from_account += i[0]
        else:
            month_to_account += i[0]

    month_balance_amount = month_to_account + month_from_account
    if(month_from_account < 0):
        month_from_account *= (-1)
        

    show = ("SELECT amount FROM accounts")
    cursor.execute(show)
    total_data=cursor.fetchall()

    total_balance_amount = 0.0
    total_from_account = 0.0
    total_to_account = 0.0
    for i in total_data:
        if(i[0] < 0):
            total_from_account += i[0]
        else:
            total_to_account += i[0]

    total_balance_amount = total_to_account + total_from_account
    if(total_from_account < 0):
        total_from_account *= (-1)

    Label(et_tk,text="This Day",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1055,y=100)
    Label(et_tk,text="From Amount : {number:06}".format(number=(today_from_account)),font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=150)
    Label(et_tk,text="To Amount : {number:06}".format(number=today_to_account),font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=190)
    Label(et_tk,text="- - - - - - - - - - - - - - - - - - - - - - -",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=230)
    Label(et_tk,text="Balance Amount : {number:06}".format(number=today_balance_amount),font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=250)
    Label(et_tk,text="- - - - - - - - - - - - - - - - - - - - - - -",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=270)

    Label(et_tk,text="This Month",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1045,y=350)
    Label(et_tk,text="From Amount : {number:06}".format(number=(month_from_account)),font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=400)
    Label(et_tk,text="To Amount : {number:06}".format(number=(month_to_account)),font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=440)
    Label(et_tk,text="- - - - - - - - - - - - - - - - - - - - - - -",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=480)
    Label(et_tk,text="Balance Amount : {number:06}".format(number=month_balance_amount),font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=500)
    Label(et_tk,text="- - - - - - - - - - - - - - - - - - - - - - -",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=520)

    Label(et_tk,text="Total",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1065,y=600)
    Label(et_tk,text="From Amount : {number:06}".format(number=(total_from_account)),font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=650)
    Label(et_tk,text="To Amount : {number:06}".format(number=total_to_account),font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=690)
    Label(et_tk,text="- - - - - - - - - - - - - - - - - - - - - - -",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=730)
    Label(et_tk,text="Balance Amount : {number:06}".format(number=total_balance_amount),font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1000,y=750)
    Label(et_tk,text="- - - - - - - - - - - - - - - - - - - - - - -",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1000,y=770)

def from_account_save():
    if(from_particulars_entry.get() == ""):
        messagebox.showerror(parent=et_tk, title="Invalid Particulars", message="Please enter a valid particulars")

    elif(from_amount_entry.get()=="" or float(from_amount_entry.get()) <= 0.0):
        messagebox.showerror(parent=et_tk, title="Invalid Amount", message="Please enter a valid amount")
    
    else:
        particulars = from_particulars_entry.get()
        amount = float(from_amount_entry.get())*(-1)
        today = date.today()
        insert = ("INSERT INTO accounts VALUES('{particulars}',{amount},'{date}')").format(date=today,particulars=particulars,amount=amount)
        cursor.execute(insert)
        mycon.commit()
    
        from_particulars_entry.delete(0,END)
        from_amount_entry.delete(0,END)
        
        display_block()
        
        messagebox.showinfo(parent=et_tk, title="Saved Data", message="Your expences data has been recorded successfully")
    
def to_account_save():
    if(to_particulars_entry.get() == ""):
        messagebox.showerror(parent=et_tk, title="Invalid Particulars", message="Please enter a valid particulars")

    elif(to_amount_entry.get()=="" or float(to_amount_entry.get()) <= 0.0):
        messagebox.showerror(parent=et_tk, title="Invalid Amount", message="Please enter a valid amount")
        
    else:
        particulars = to_particulars_entry.get()
        amount = float(to_amount_entry.get())
        today = date.today()
        insert = ("INSERT INTO accounts VALUES('{particulars}',{amount},'{date}')").format(date=today,particulars=particulars,amount=amount)
        cursor.execute(insert)
        mycon.commit()
    
        to_particulars_entry.delete(0,END)
        to_amount_entry.delete(0,END)
        
        display_block()
        
        messagebox.showinfo(parent=et_tk, title="Saved Data", message="Your expences data has been recorded successfully")

def generate_report():
    date1 = str(from_date_entry.get())
    date2 = str(to_date_entry.get())
    
    if(date1 == ""):
        messagebox.showerror(parent=et_tk, title="Invalid Date", message="Please enter a valid from date")
    
    elif(date2 == ""):
        messagebox.showerror(parent=et_tk, title="Invalid Date", message="Please enter a valid to date")
    
    elif(date2 < date1):
        messagebox.showerror(parent=et_tk, title="Invalid Date", message="To date should be greater than from date")
    
    elif(date2 > date.today().strftime("%Y-%m-%d")):
        messagebox.showerror(parent=et_tk, title="Invalid Date", message="To date should be less than or equal to today's date")
    
    else:
        show = ("SELECT * FROM accounts where Date BETWEEN '{}' AND '{}'").format(date1,date2)
        cursor.execute(show)
        data=cursor.fetchall()
        
        if(len(data) == 0):
            messagebox.showerror(parent=et_tk, title="No Data", message="No data found for the given date range")
            return    
        
        try:
            show = ("SELECT SUM(amount) FROM accounts where Date < '{}' ").format(date1)
            cursor.execute(show)
            open_balance=float(cursor.fetchall()[0][0])
        except:
            open_balance = 0.0
        
        try:
            show = ("SELECT SUM(amount) FROM accounts where Date <= '{}' ").format(date2)
            cursor.execute(show)
            close_balance=float(cursor.fetchall()[0][0])
        except:
            close_balance = 0.0
            
        data_list = []
        from_total = 0.0
        to_total = 0.0
        for i in data:
            nest_list = []
            nest_list.append(i[2])
            nest_list.append(i[0])
        
            if(float(i[1]) < 0):
                from_total += float(i[1])*(-1)
                nest_list.append(i[1]*(-1))
                nest_list.append("")
            else:
                to_total += float(i[1])
                nest_list.append("")
                nest_list.append(i[1])
        
            data_list.append(nest_list)
    
        doc = DocxTemplate("expence_template.docx")
    
        doc.render({"from_date":date1,
                "to_date":date2,
                "expence_list": data_list,
                "total_from":from_total,
                "total_to":to_total,
                "open_balance":open_balance,
                "close_balance":close_balance})
    
        doc_name = "Expence_Report_"+date1+"_"+date2+".docx"
        doc.save('F:/ABC Project/Consultancy Project/Project/Expences Report/'+doc_name)
        
        messagebox.showinfo(parent=et_tk,title="Report Saved", message="Report Saved Successfully")
    
Label(et_tk,text="AJRA TEX - KARUR",font=("Arial", 20, "bold"),bg="white",fg="#1A374D").place(x=650,y=20)

'''------- From Account -------'''

Label(et_tk, text="- - - - - From Account - - - - -",font=("Arial", 13,"bold"),bg="white",fg="#1A374D").place(x=470,y=100)
from_particulars_lable = Label(et_tk, text="Particulars",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=405,y=150)
from_particulars_entry = Entry(et_tk,font=("Arial", 12),bd=3)
from_particulars_entry.place(x=350,y=180)

from_amount_lable = Label(et_tk, text="Amount",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=655,y=150)
from_amount_entry = Entry(et_tk,font=("Arial", 12),bd=3)
from_amount_entry.place(x=600,y=180)

from_save_button = Button(et_tk, text="Save",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",width=50,command=from_account_save)
from_save_button.place(x=370,y=230)

'''------- To Account -------'''

Label(et_tk, text="- - - - - To Account - - - - -",font=("Arial", 13,"bold"),bg="white",fg="#1A374D").place(x=480,y=350)
to_particulars_lable = Label(et_tk, text="Particulars",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=405,y=400)
to_particulars_entry = Entry(et_tk,font=("Arial", 12),bd=3)
to_particulars_entry.place(x=350,y=430)

to_amount_lable = Label(et_tk, text="Amount",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=655,y=400)
to_amount_entry = Entry(et_tk,font=("Arial", 12),bd=3)
to_amount_entry.place(x=600,y=430)

to_save_button = Button(et_tk, text="Save",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",width=50,command=to_account_save)
to_save_button.place(x=370,y=480)

'''------- Generate Report -------'''

Label(et_tk, text="- - - - - Generate Report - - - - -",font=("Arial", 13,"bold"),bg="white",fg="#1A374D").place(x=460,y=600)
from_date_lable = Label(et_tk, text="From Date",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=400,y=650)
from_date_entry = Entry(et_tk,font=("Arial", 12),bd=3)
from_date_entry.place(x=350,y=680)

to_date_lable = Label(et_tk, text="To Date",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=655,y=630)
to_date_entry = Entry(et_tk,font=("Arial", 12),bd=3)
to_date_entry.place(x=600,y=680)

generate_button = Button(et_tk, text="Generate Report",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",width=50,command=generate_report)
generate_button.place(x=370,y=730)

display_block()

et_tk.mainloop()