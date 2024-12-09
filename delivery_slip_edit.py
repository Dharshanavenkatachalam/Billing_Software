from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
import time
from tkinter import messagebox
import mysql.connector as sql
import win32print

ds_edit_tk=Tk()
ds_edit_tk.geometry("1920x1080")
ds_edit_tk.config(bg="#F5F5F5")
ds_edit_tk.title("Delivery Slip Edit - Ajra Tex")
ds_edit_tk.configure(bg='white')

mycon=sql.connect(host="localhost",user="root",passwd="gokul123",database="ajra")
cursor=mycon.cursor()

def clear_item():
    particulars_entry.delete(0,END)
    quantity_entry.delete(0,END)
    quantity_entry.insert(0, "0")
        
def add_item():
    tree_items = tree.get_children()
    invoice_list = []
    
    for i in tree_items:
        item = tree.item(i)['values']
        invoice_list.append(item[0].lower())
        
    if(particulars_entry.get() == ""):
        messagebox.showerror(parent=ds_edit_tk, title="Invalid Product Name", message="Please enter a valid product name")
        return
    
    elif(int(quantity_entry.get()) < 1):
        messagebox.showerror(parent=ds_edit_tk, title="Invalid Quantity", message="Please enter a valid quantity count")
        return
    
    elif(len(tree.selection())<1):
        if((particulars_entry.get()).lower() in invoice_list):
            messagebox.showerror(parent=ds_edit_tk, title="Duplicate Product", message="Product already exists in the delivery slip")
            return
        particulars = particulars_entry.get()
        quantity = int(quantity_entry.get())
        invoice_item = [particulars, quantity]
        tree.insert('',0, values=invoice_item)
        clear_item()
            
    else:
        selected_item = tree.selection()[0]  
        tree.delete(selected_item)
                
        particulars = particulars_entry.get()
        quantity = int(quantity_entry.get())
        invoice_item = [particulars, quantity]
        tree.insert('',0, values=invoice_item)
        clear_item()
    
def edit_invoice():
    try:
        clear_entry()
        inv_no = slip_number_entry.get()
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
        messagebox.showerror(parent=ds_edit_tk, title="Invalid Delivery Slip Number", message="Please enter a valid delivery slip number")
        
def edit_selection():
    if(len(tree.selection())<1):
        messagebox.showerror(parent=ds_edit_tk, title="Selection Error", message="Please select an item to edit")
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
        messagebox.showerror(parent=ds_edit_tk, title="Selection Error", message="Please select an item to delete")
        return
    
    else:
        answer = messagebox.askyesno(parent=ds_edit_tk, title="Delete Item", message="Are you sure to delete the selected item?")
        
        if(answer == True):
            selected_item = tree.selection()[0]
            tree.delete(selected_item)

def clear_entry():
    name_entry.delete(0,END)
    address_entry.delete(0,END)
    phone_number_entry.delete(0,END)
    gst_entry.delete(0,END)
    
    particulars_entry.delete(0,END)
    quantity_entry.delete(0,END)
    quantity_entry.insert(0, "0")
    
    for item in tree.get_children():
        tree.delete(item)

def new_deliver_slip():
    answer = messagebox.askyesno(parent=ds_edit_tk, title="New Delivery Slip", message="Do you want to create a new delivery slip?")
    
    if(answer== True):
        name_entry.delete(0,END)
        address_entry.delete(0,END)
        phone_number_entry.delete(0,END)
        gst_entry.delete(0,END)
    
        slip_number_entry.delete(0,END)
    
        particulars_entry.delete(0,END)
        quantity_entry.delete(0,END)
        quantity_entry.insert(0, "0")
    
        for item in tree.get_children():
            tree.delete(item)

def generate_delivery_slip():
    if(name_entry.get() == ""):
        messagebox.showerror(parent=ds_edit_tk, title="Invalid Name", message="Please enter a valid name")
        return
    
    elif(address_entry.get() == ""):
        messagebox.showerror(parent=ds_edit_tk, title="Invalid Address", message="Please enter a valid address")
        return
    
    elif(phone_number_entry.get() == "" or len(phone_number_entry.get())!=10):
        messagebox.showerror(parent=ds_edit_tk, title="Invalid Phone Number", message="Please enter a valid phone number")
        return
    
    elif(gst_entry.get() == "" or len(gst_entry.get())!=15):
        messagebox.showerror(parent=ds_edit_tk, title="Invalid GST Number", message="Please enter a valid GST number")
        return 
    
    elif(len(tree.get_children()) == 0):
        messagebox.showerror(parent=ds_edit_tk, title="Empty Invoice", message="Please add items to the invoice")
        return
    
    else:
        doc = DocxTemplate("delivery_slip_template.docx")
        name = name_entry.get()
        address = address_entry.get()
        phone = phone_number_entry.get()
        gst = gst_entry.get()
        dsno = slip_number_entry.get()
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
    
        doc_name = slip_number_entry.get()+".docx"
        doc.save('F:/ABC Project/Consultancy Project/Project/Delivery Slips/'+doc_name)
    
        particulars_data = ""
        quantity_data = ""
    
        for i in range(len(delivery_list)):
            particulars_data += delivery_list[i][0]
            quantity_data += str(delivery_list[i][1])
                
            if(i<len(delivery_list)-1):
                particulars_data += ","
                quantity_data += ","
                
        update = ("UPDATE delivery_slip SET delivery_slip_date='{}', delivery_slip_time='{}', name='{}', address='{}', phone_no='{}', gst_no='{}', particulars='{}', quantity='{}' WHERE delivery_slip_no='{}'").format(current_date,current_time,name,address,phone,gst,particulars_data,quantity_data,dsno)
        cursor.execute(update)
        mycon.commit()
    
        answer = messagebox.askyesno(parent=ds_edit_tk, title="Delivery Slip", message="Do you want to print the delivery slip?")
    
        if(answer == True):
            print_document()
    
        messagebox.showinfo(parent=ds_edit_tk, title="Delivery Slip", message="Delivery Slip Generated Successfully")

def print_document():
    file_path = "F:/ABC Project/Consultancy Project/Project/Delivery Slips/"+slip_number_entry.get()+".docx"
    printer_path = 'Microsoft Print to PDF'
    file_handle = open(file_path, 'rb')
    
    printer_handle = win32print.OpenPrinter(printer_path)
    JobInfo = win32print.StartDocPrinter(printer_handle, 1, (file_path, None, "RAW"))
    win32print.StartPagePrinter(printer_handle)
    win32print.WritePrinter(printer_handle, file_handle.read())
    win32print.EndPagePrinter(printer_handle)
   
Label(ds_edit_tk,text="AJRA TEX - KARUR",font=("Arial", 20, "bold"),bg="white",fg="#1A374D").place(x=650,y=20)
Label(ds_edit_tk,text="Delivery Number : ",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=580,y=100)

slip_number_entry = Entry(ds_edit_tk,font=("Arial", 12, "bold"),bd=3)
slip_number_entry.place(x=730,y=100)

edit_button = Button(ds_edit_tk, text="Submit",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",command=edit_invoice)
edit_button.place(x=930,y=97)

name_label = Label(ds_edit_tk, text="Name",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=260,y=150)
name_entry = Entry(ds_edit_tk,font=("Arial", 12),bd=3)
name_entry.place(x=190,y=180)

address_label = Label(ds_edit_tk, text="Address",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=560,y=150)
address_entry = Entry(ds_edit_tk,font=("Arial", 12),bd=3)
address_entry.place(x=500,y=180)

phone_number_label = Label(ds_edit_tk, text="Phone number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=850,y=150)
phone_number_entry = Entry(ds_edit_tk,font=("Arial", 12),bd=3)
phone_number_entry.place(x=810,y=180)

gst_label = Label(ds_edit_tk, text="GST number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1160,y=150)
gst_entry = Entry(ds_edit_tk,font=("Arial", 12),bd=3)
gst_entry.place(x=1110,y=180)

'''------- Product details -------'''

particulars_label = Label(ds_edit_tk, text="Particulars",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=550,y=230)
particulars_entry = Entry(ds_edit_tk,font=("Arial", 12),bd=3)
particulars_entry.place(x=500,y=260)

quantity_label = Label(ds_edit_tk, text="Quantity",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=870,y=230)
quantity_entry = Spinbox(ds_edit_tk, from_=0,to=100,font=("Arial", 12),bd=3)
quantity_entry.place(x=810,y=260)

add_item_button = Button(ds_edit_tk, text="Add item",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",command=add_item)
add_item_button.place(x=1110,y=260)

columns = ('particulars', 'quantity')
tree = ttk.Treeview(ds_edit_tk, columns=columns, show="headings")
tree.heading('particulars', text='Particulars')
tree.heading('quantity', text='Quantity')
tree.place(x=580,y=330)

edit_btn = ttk.Button(ds_edit_tk,text="Edit", command=edit_selection)
edit_btn.place(x=900,y=570)

delete_btn = ttk.Button(ds_edit_tk,text="Delete",command=delete_selection)
delete_btn.place(x=815,y=570)

save_invoice_button = Button(ds_edit_tk, text="Generate Delivery Slip",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=generate_delivery_slip)
save_invoice_button.place(x=650,y=640)

new_invoice_button = Button(ds_edit_tk, text="New Delivery Slip",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=new_deliver_slip)
new_invoice_button.place(x=650,y=700)

ds_edit_tk.mainloop()
