from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import mysql.connector as sql
import win32print

ds_view=Tk()
ds_view.geometry("1920x1080")
ds_view.config(bg="#F5F5F5")
ds_view.title("View Delivery Slip - Ajra Tex")
ds_view.configure(bg='white')

mycon=sql.connect(host="localhost",user="root",passwd="gokul123",database="ajra")
cursor=mycon.cursor()
        
def view_invoice():
    clear_invoice()
    try:
        inv_no = str(slip_number_entry.get())
        show = ("SELECT * FROM delivery_slip WHERE delivery_slip_no = '{}'").format(inv_no)
        cursor.execute(show)
        data=cursor.fetchall()[0]
        
        name_entry.insert(0,data[3])
        address_entry.insert(0,data[4])
        phone_number_entry.insert(0,data[5])
        gst_entry.insert(0,data[6])
   
        particulars = data[7].split(',')
        quantity = data[8].split(',')
    
        for i in range(len(particulars)):
            invoice_item = [particulars[i], int(quantity[i])]
            tree.insert('',0, values=invoice_item)
    
    except:
         messagebox.showerror(parent=ds_view, title="Invalid Delivery Slip Number", message="Please enter a valid delivery slip number")
        
def print_document():
    answer = messagebox.askyesno(parent=ds_view, title="Print delivery slip", message="Are you sure to print the delivery slip?")
    
    if(answer == True):
        try:
            file_path = "F:/ABC Project/Consultancy Project/Project/Delivery Slips/"+slip_number_entry.get()+".docx"
            printer_path = 'Microsoft Print to PDF'
            file_handle = open(file_path, 'rb')
    
            printer_handle = win32print.OpenPrinter(printer_path)
            JobInfo = win32print.StartDocPrinter(printer_handle, 1, (file_path, None, "RAW"))
            win32print.StartPagePrinter(printer_handle)
            win32print.WritePrinter(printer_handle, file_handle.read())
            win32print.EndPagePrinter(printer_handle)
        except:
            messagebox.showerror(parent=ds_view, title="Empty Field Error", message="Enter the invoice number to print the invoice")

def new_invoice():
    answer=messagebox.askyesno("Clear Delivery Slip","Are you sure to clear the delivery slip?")
    
    if(answer==True):
        name_entry.delete(0,END)
        address_entry.delete(0,END)
        phone_number_entry.delete(0,END)
        gst_entry.delete(0,END)
    
        slip_number_entry.delete(0,END)
            
        for item in tree.get_children():
            tree.delete(item)

def clear_invoice():
    name_entry.delete(0,END)
    address_entry.delete(0,END)
    phone_number_entry.delete(0,END)
    gst_entry.delete(0,END)
    
    for item in tree.get_children():
        tree.delete(item)
                
Label(ds_view,text="AJRA TEX - KARUR",font=("Arial", 20, "bold"),bg="white",fg="#1A374D").place(x=650,y=20)
Label(ds_view,text="Invoice Number : ",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=580,y=100)

slip_number_entry = Entry(ds_view,font=("Arial", 12, "bold"),bd=3)
slip_number_entry.place(x=730,y=100)

edit_button = Button(ds_view, text="Submit",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",command=view_invoice)
edit_button.place(x=930,y=97)

name_label = Label(ds_view, text="Name",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=260,y=170)
name_entry = Entry(ds_view,font=("Arial", 12),bd=3)
name_entry.place(x=190,y=200)

address_label = Label(ds_view, text="Address",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=560,y=170)
address_entry = Entry(ds_view,font=("Arial", 12),bd=3)
address_entry.place(x=500,y=200)

phone_number_label = Label(ds_view, text="Phone number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=850,y=170)
phone_number_entry = Entry(ds_view,font=("Arial", 12),bd=3)
phone_number_entry.place(x=810,y=200)

gst_label = Label(ds_view, text="GST number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1160,y=170)
gst_entry = Entry(ds_view,font=("Arial", 12),bd=3)
gst_entry.place(x=1110,y=200)

columns = ('particulars', 'quantity')
tree = ttk.Treeview(ds_view, columns=columns, show="headings")
tree.heading('particulars', text='Particulars')
tree.heading('quantity', text='Quantity')
tree.place(x=560,y=320)

save_invoice_button = Button(ds_view, text="Print Delivery Slip",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=print_document)
save_invoice_button.place(x=650,y=600)

clear_invoice_button = Button(ds_view, text="Clear Delivery Slip",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=new_invoice)
clear_invoice_button.place(x=650,y=680)

ds_view.mainloop()
