from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import mysql.connector as sql
import win32print

in_view=Tk()
in_view.geometry("1920x1080")
in_view.config(bg="#F5F5F5")
in_view.title("View Invoice - Ajra Tex")
in_view.configure(bg='white')

mycon=sql.connect(host="localhost",user="root",passwd="gokul123",database="ajra")
cursor=mycon.cursor()

def clear_invoice():
    name_entry.delete(0,END)
    address_entry.delete(0,END)
    phone_number_entry.delete(0,END)
    gst_entry.delete(0,END)
    
    stotal_entry.delete(0,END)
    stotal_entry.insert(0,"0.0")
    cgst_entry.delete(0,END)
    cgst_entry.insert(0,"0.0")
    sgst_entry.delete(0,END)
    sgst_entry.insert(0,"0.0")
    gtotal_entry.delete(0,END)
    gtotal_entry.insert(0,"0.0")
    
    for item in tree.get_children():
        tree.delete(item)   

def new_invoice():
    answer=messagebox.askyesno(parent=in_view, title="Clear Invoice", message="Are you sure to clear the invoice?")
    
    if(answer==True):
        name_entry.delete(0,END)
        address_entry.delete(0,END)
        phone_number_entry.delete(0,END)
        gst_entry.delete(0,END)
    
        invoice_number_entry.delete(0,END)
        
        stotal_entry.delete(0,END)
        stotal_entry.insert(0,"0.0")
        cgst_entry.delete(0,END)
        cgst_entry.insert(0,"0.0")
        sgst_entry.delete(0,END)
        sgst_entry.insert(0,"0.0")
        gtotal_entry.delete(0,END)
        gtotal_entry.insert(0,"0.0")
    
        for item in tree.get_children():
            tree.delete(item)   

def view_invoice():
    clear_invoice()
    try:
        inv_no = str(invoice_number_entry.get())
        show = ("SELECT * FROM invoice WHERE invoice_no = '{}'").format(inv_no)
        cursor.execute(show)
        data=cursor.fetchall()[0]
        
        name_entry.insert(0,data[3])
        address_entry.insert(0,data[4])
        phone_number_entry.insert(0,data[5])
        gst_entry.insert(0,data[6])
    
        stotal_entry.delete(0,END)
        stotal_entry.insert(0,data[11])
        
        cgst_entry.delete(0,END)
        cgst_entry.insert(0,data[12])
        
        sgst_entry.delete(0,END)
        sgst_entry.insert(0,data[13])
        
        gtotal_entry.delete(0,END)
        gtotal_entry.insert(0,data[14])
    
        particulars = data[7].split(',')
        quantity = data[8].split(',')
        unit_price = data[9].split(',')
        line_total = data[10].split(',')
    
        for i in range(len(particulars)):
            invoice_item = [particulars[i], int(quantity[i]), float(unit_price[i]), float(line_total[i])]
            tree.insert('',0, values=invoice_item)
    
    except:
         messagebox.showerror(parent=in_view, title="Invalid Invoice Number", message="Please enter a valid invoice number")
        
def print_document():
    answer = messagebox.askyesno(parent=in_view, title="Print invoice", message="Are you sure to print the invoice?")
    
    if(answer == True):
        try:
            file_path = "F:/ABC Project/Consultancy Project/Project/Invoices/"+invoice_number_entry.get()+".docx"
            printer_path = 'Microsoft Print to PDF'
            file_handle = open(file_path, 'rb')
    
            printer_handle = win32print.OpenPrinter(printer_path)
            JobInfo = win32print.StartDocPrinter(printer_handle, 1, (file_path, None, "RAW"))
            win32print.StartPagePrinter(printer_handle)
            win32print.WritePrinter(printer_handle, file_handle.read())
            win32print.EndPagePrinter(printer_handle)
        except:
            messagebox.showerror(parent=in_view, title="Empty Field Error", message="Enter the invoice number to print the invoice")
    
Label(in_view,text="AJRA TEX - KARUR",font=("Arial", 20, "bold"),bg="white",fg="#1A374D").place(x=650,y=20)
Label(in_view,text="Invoice Number : ",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=580,y=100)

invoice_number_entry = Entry(in_view,font=("Arial", 12, "bold"),bd=3)
invoice_number_entry.place(x=730,y=100)

edit_button = Button(in_view, text="Submit",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",command=view_invoice)
edit_button.place(x=930,y=97)

name_label = Label(in_view, text="Name",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=260,y=170)
name_entry = Entry(in_view,font=("Arial", 12),bd=3)
name_entry.place(x=190,y=200)

address_label = Label(in_view, text="Address",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=560,y=170)
address_entry = Entry(in_view,font=("Arial", 12),bd=3)
address_entry.place(x=500,y=200)

phone_number_label = Label(in_view, text="Phone number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=850,y=170)
phone_number_entry = Entry(in_view,font=("Arial", 12),bd=3)
phone_number_entry.place(x=810,y=200)

gst_label = Label(in_view, text="GST number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1160,y=170)
gst_entry = Entry(in_view,font=("Arial", 12),bd=3)
gst_entry.place(x=1110,y=200)

columns = ('particulars', 'quantity', 'unit_price', 'amount')
tree = ttk.Treeview(in_view, columns=columns, show="headings")
tree.heading('particulars', text='Particulars')
tree.heading('quantity', text='Quantity')
tree.heading('unit_price', text='Unit Price')
tree.heading('amount', text="Amount")
tree.place(x=190,y=320)

stotal_label = Label(in_view, text="Sub Total : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=340)
stotal_entry = Entry(in_view,font=("Arial", 12),bd=3)
stotal_entry.place(x=1220,y=340)
stotal_entry.insert(0,"0.0")

cgst_label = Label(in_view, text="CGST @2.5% : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=390)
cgst_entry = Entry(in_view,font=("Arial", 12),bd=3)
cgst_entry.place(x=1220,y=390)
cgst_entry.insert(0,"0.0")

sgst_label = Label(in_view, text="SGST @2.5% : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=440)
sgst_entry = Entry(in_view,font=("Arial", 12),bd=3)
sgst_entry.place(x=1220,y=440)
sgst_entry.insert(0,"0.0")

gtotal_label = Label(in_view, text="Grand Total : ",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1100,y=490)
gtotal_entry = Entry(in_view,font=("Arial", 12, "bold"),bd=3)
gtotal_entry.place(x=1220,y=490)
gtotal_entry.insert(0,"0.0")

save_invoice_button = Button(in_view, text="Print Invoice",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=print_document)
save_invoice_button.place(x=650,y=600)

clear_invoice_button = Button(in_view, text="Clear Invoice",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=new_invoice)
clear_invoice_button.place(x=650,y=680)

in_view.mainloop()
