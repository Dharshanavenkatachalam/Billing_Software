from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
import time
from tkinter import messagebox
import mysql.connector as sql
import win32print
import os

ds_tk=Tk()
ds_tk.geometry("1920x1080")
ds_tk.config(bg="#F5F5F5")
ds_tk.title("Delivery Slip - Ajra Tex")
ds_tk.configure(bg='white')

mycon=sql.connect(host="localhost",user="root",passwd="gokul123",database="ajra")
cursor=mycon.cursor()

show = ("SELECT COUNT(*) FROM delivery_slip")
cursor.execute(show)
slip_no=cursor.fetchall()
slip_no=int(slip_no[0][0])+1

def clear_item():
    particulars_entry.delete(0,END)
    quantity_entry.delete(0,END)
    quantity_entry.insert(0, "0")
    
def add_item():
    particulars = particulars_entry.get()
    quantity = int(quantity_entry.get())
    
    tree_items = tree.get_children()
    invoice_list = []
    
    for i in tree_items:
        item = tree.item(i)['values']
        invoice_list.append(item[0].lower())
        
    if(particulars == ""):
        messagebox.showerror(parent=ds_tk, title="Invalid Product Name", message="Please enter a valid product name")
        return
    
    elif(quantity < 1):
        messagebox.showerror(parent=ds_tk, title="Invalid Quantity", message="Please enter a valid quantity count")
        return
    
    elif(len(tree.selection())<1):
        if(particulars.lower() in invoice_list):
            messagebox.showerror(parent=ds_tk, title="Duplicate Product", message="Product already exists in the delivery slip")
            return
        else:
            particulars = particulars_entry.get()
            quantity = int(quantity_entry.get())
            delivery_item = [particulars, quantity]
            tree.insert('',0, values=delivery_item)
            clear_item()
            
    else:
        selected_item = tree.selection()[0]  
        tree.delete(selected_item)
                
        particulars = particulars_entry.get()
        quantity = int(quantity_entry.get())
        delivery_item = [particulars, quantity]
        tree.insert('',0, values=delivery_item)
        clear_item()
        
def generate_delivery_slip():
    try:
        if(name_entry.get() == ""):
            messagebox.showerror(parent=ds_tk, title="Invalid Name", message="Please enter a valid name")
            return
    
        elif(address_entry.get() == ""):
            messagebox.showerror(parent=ds_tk, title="Invalid Address", message="Please enter a valid address")
            return
    
        elif(phone_number_entry.get() == "" or len(phone_number_entry.get())!=10):
            messagebox.showerror(parent=ds_tk, title="Invalid Phone Number", message="Please enter a valid phone number")
            return
    
        elif(gst_entry.get() == "" or len(gst_entry.get())!=15):
            messagebox.showerror(parent=ds_tk, title="Invalid GST Number", message="Please enter a valid GST number")
            return 
    
        elif(len(tree.get_children()) == 0):
            messagebox.showerror(parent=ds_tk, title="Empty Invoice", message="Please add items to the invoice")
            return
    
        else:
            doc = DocxTemplate("delivery_slip_template.docx")
            name = name_entry.get()
            address = address_entry.get()
            phone = phone_number_entry.get()
            gst = gst_entry.get()
            dsno = "DS-2425-{number:06}".format(number=slip_no)
            current_date = datetime.date.today()
            t = time.localtime()
            current_time = time.strftime("%H:%M:%S", t)
            
            delivery_list = []
            tree_items = tree.get_children()
    
            for i in tree_items:
                item = tree.item(i)['values']
                delivery_list.append(item)
                
            doc.render({"name":name,
                "address":address,
                "phone":phone,
                "gst":gst,
                "deliveryno": dsno,
                "deliverydate":current_date,
                "deliverytime":current_time,
                "delivery_list": delivery_list})
    
            doc_name = "DS-2425-{number:06}".format(number=slip_no)+".docx"
            doc.save('F:/ABC Project/Consultancy Project/Project/Delivery Slips/'+doc_name)
    
            particulars_data = ""
            quantity_data = ""
    
            for i in range(len(delivery_list)):
                particulars_data += delivery_list[i][0]
                quantity_data += str(delivery_list[i][1])
                
                if(i<len(delivery_list)-1):
                    particulars_data += ","
                    quantity_data += ","
                
            insert = ("INSERT INTO delivery_slip VALUES('{}','{}','{}','{}','{}','{}','{}','{}','{}')").format(dsno,current_date,current_time,name,address,phone,gst,particulars_data,quantity_data)
            cursor.execute(insert)
            mycon.commit()
            
            answer = messagebox.askyesno(parent=ds_tk, title="Delivery Slip", message="Do you want to print the delivery slip?")
    
            if(answer == True):
                print_document()
    
            messagebox.showinfo(parent=ds_tk, title="Delivery Slip", message="Delivery Slip Generated Successfully")
        
    except:
        messagebox.showerror(parent=ds_tk, title="Duplicate Delivery Slip", message="Delivery slip already exists in the database. Please create a new delivery slip.")
    
def clear_delivery_slip():
    answer = messagebox.askyesno(parent=ds_tk,title="New Delivery Slip", message="Do you want to create a new delivery slip?")
    
    if(answer == True):
        name_entry.delete(0,END)
        address_entry.delete(0,END)
        phone_number_entry.delete(0,END)
        gst_entry.delete(0,END)
    
        particulars_entry.delete(0,END)
        quantity_entry.delete(0,END)
        quantity_entry.insert(0, "0")
        
        for item in tree.get_children():
            tree.delete(item)
        
        show = ("SELECT COUNT(*) FROM delivery_slip")
        cursor.execute(show)
        slip_no=cursor.fetchall()
        slip_no=int(slip_no[0][0])+1
        
        Label(ds_tk,text="Delivery Slip Number : DS-2425-{number:06}".format(number=slip_no),font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1200,y=20)

def edit_selection():
    if(len(tree.selection())<1):
        messagebox.showerror(parent=ds_tk, title="Selection Error", message="Please select an item to edit")
        return
    
    else:
        selected_item = tree.selection()[0]
        item = tree.item(selected_item)['values']
               
        particulars_entry.delete(0,END)
        particulars_entry.insert(0,item[0])
        quantity_entry.delete(0,END)
        quantity_entry.insert(0,int(item[1]))

def delete_selection():
    if(len(tree.selection())<1):
        messagebox.showerror(parent=ds_tk, title="Selection Error", message="Please select an item to delete")
        return
    
    else:
        answer = messagebox.askyesno(parent=ds_tk, title="Delete Item", message="Are you sure to delete the selected item?")
        
        if(answer == True):
            selected_item = tree.selection()[0]
            tree.delete(selected_item)

def print_document():
    file_path = "F:/ABC Project/Consultancy Project/Project/Delivery Slips/"+"DS-2425-{number:06}".format(number=slip_no)+".docx"
    printer_path = 'Microsoft Print to PDF'
    file_handle = open(file_path, 'rb')
    
    printer_handle = win32print.OpenPrinter(printer_path)
    JobInfo = win32print.StartDocPrinter(printer_handle, 1, (file_path, None, "RAW"))
    win32print.StartPagePrinter(printer_handle)
    win32print.WritePrinter(printer_handle, file_handle.read())
    win32print.EndPagePrinter(printer_handle)

def edit_delivery_slip():
    os.system('python delivery_slip_edit.py')

def view_delivery_slip():
    os.system('python delivery_slip_view.py')
    

Label(ds_tk,text="AJRA TEX - KARUR",font=("Arial", 20, "bold"),bg="white",fg="#1A374D").place(x=650,y=20)
Label(ds_tk,text="Delivery Slip Number : DS-2425-{number:06}".format(number=slip_no),font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1200,y=20)

'''------- Buyer details -------'''

name_label = Label(ds_tk, text="Name",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=260,y=100)
name_entry = Entry(ds_tk,font=("Arial", 12),bd=3)
name_entry.place(x=190,y=130)

address_label = Label(ds_tk, text="Address",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=560,y=100)
address_entry = Entry(ds_tk,font=("Arial", 12),bd=3)
address_entry.place(x=500,y=130)

phone_number_label = Label(ds_tk, text="Phone number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=850,y=100)
phone_number_entry = Entry(ds_tk,font=("Arial", 12),bd=3)
phone_number_entry.place(x=810,y=130)

gst_label = Label(ds_tk, text="GST number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1160,y=100)
gst_entry = Entry(ds_tk,font=("Arial", 12),bd=3)
gst_entry.place(x=1110,y=130)

'''------- Product details -------'''

particulars_label = Label(ds_tk, text="Particulars",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=550,y=180)
particulars_entry = Entry(ds_tk,font=("Arial", 12),bd=3)
particulars_entry.place(x=500,y=210)

quantity_label = Label(ds_tk, text="Quantity",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=870,y=180)
quantity_entry = Spinbox(ds_tk, from_=0,to=100,font=("Arial", 12),bd=3)
quantity_entry.place(x=810,y=210)

add_item_button = Button(ds_tk, text="Add item",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",command=add_item)
add_item_button.place(x=1110,y=210)

edit_btn = ttk.Button(ds_tk,text="Edit", command=edit_selection)
edit_btn.place(x=900,y=535)

delete_btn = ttk.Button(ds_tk,text="Delete",command=delete_selection)
delete_btn.place(x=815,y=535)

columns = ('particulars', 'quantity')
tree = ttk.Treeview(ds_tk, columns=columns, show="headings")
tree.heading('particulars', text='Particulars')
tree.heading('quantity', text='Quantity')
tree.place(x=580,y=300)

save_slip_button = Button(ds_tk, text="Generate Delivery Slip",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=generate_delivery_slip)
save_slip_button.place(x=650,y=580)

new_slip_button = Button(ds_tk, text="New Delivery Slip",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=clear_delivery_slip)
new_slip_button.place(x=650,y=640)

edit_slip_button = Button(ds_tk, text="Edit Delivery Slip",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=edit_delivery_slip)
edit_slip_button.place(x=650,y=700)

view_slip_button = Button(ds_tk, text="View Delivery Slip",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=view_delivery_slip)
view_slip_button.place(x=650,y=760)

ds_tk.mainloop()
