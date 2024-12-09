from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
from datetime import date
import time
from tkinter import messagebox
import mysql.connector as sql
import win32print
import os
import inflect

invoice_tk=Tk()
invoice_tk.geometry("1920x1080")
invoice_tk.config(bg="#F5F5F5")
invoice_tk.title("Invoice - Ajra Tex")
invoice_tk.configure(bg='white')

mycon=sql.connect(host="localhost",user="root",passwd="gokul123",database="ajra")
cursor=mycon.cursor()

show = ("SELECT COUNT(*) FROM invoice")
cursor.execute(show)
invoice_no=cursor.fetchall()
invoice_no=int(invoice_no[0][0])+1

def number_to_words_inr(number):
    def num_to_words(num):
        below_20 = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", 
                    "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
        tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
        if num < 20:
            return below_20[num]
        elif num < 100:
            return tens[num // 10] + (" " + below_20[num % 10] if num % 10 != 0 else "")
        elif num < 1000:
            return below_20[num // 100] + " Hundred" + (" and " + num_to_words(num % 100) if num % 100 != 0 else "")
        else:
            return ""

    def convert_to_indian_numbering(num):
        units = ["", "Thousand", "Lakhs", "Crore"]
        parts = []
        units_idx = 0
        
        while num > 0:
            if units_idx == 0:
                parts.append((num % 1000, units[units_idx]))
                num //= 1000
            else:
                parts.append((num % 100, units[units_idx]))
                num //= 100
            units_idx += 1
            
        parts = [(value, unit) for value, unit in parts if value > 0]
        
        words = []
        for value, unit in reversed(parts):
            if unit:
                words.append(num_to_words(value) + " " + unit)
            else:
                words.append(num_to_words(value))
        
        return " ".join(words).strip()

    if isinstance(number, float):
        integer_part = int(number)
        fractional_part = round((number - integer_part) * 100)
    else:
        integer_part = number
        fractional_part = 0

    integer_part_in_words = convert_to_indian_numbering(integer_part)
    
    if fractional_part > 0:
        fractional_part_in_words = num_to_words(fractional_part)
        return f"{integer_part_in_words} rupees and {fractional_part_in_words} paise only."
    else:
        return f"{integer_part_in_words} rupees only."
    
def clear_item():
    particulars_entry.delete(0,END)
    quantity_entry.delete(0,END)
    quantity_entry.insert(0, "0")
    price_entry.delete(0,END)
    price_entry.insert(0, "0.0")

def amount_calculator():
    tree_items = tree.get_children()
    invoice_list = []
    
    for i in tree_items:
        item = tree.item(i)['values']
        invoice_list.append(item)
            
    sub_total = sum(float(item[3]) for item in invoice_list)
    stotal_entry.delete(0,END)
    stotal_entry.insert(0,sub_total)
    
    cgst_entry.delete(0,END)
    cgst_entry.insert(0,sub_total*0.025)
    
    sgst_entry.delete(0,END)
    sgst_entry.insert(0,sub_total*0.025)
    
    gtotal_entry.delete(0,END)
    gtotal_entry.insert(0,sub_total+sub_total*0.05)

def delete_selection():
    if(len(tree.selection())<1):
        messagebox.showerror(parent=invoice_tk, title="Selection Error", message="Please select an item to edit")
        return
    
    else:
        answer = messagebox.askyesno(parent=invoice_tk, title="Delete Item", message="Are you sure to delete the selected item?")
        
        if(answer == True):
            selected_item = tree.selection()[0]
            tree.delete(selected_item)
            
            amount_calculator()
                        
def edit_selection():
    if(len(tree.selection())<1):
        messagebox.showerror("Selection Error", "Please select an item to edit")
        return
    
    else:
        selected_item = tree.selection()[0]
        item = tree.item(selected_item)['values']
               
        particulars_entry.delete(0,END)
        particulars_entry.insert(0,item[0])
        quantity_entry.delete(0,END)
        quantity_entry.insert(0,int(item[1]))
        price_entry.delete(0,END)
        price_entry.insert(0,float(item[2]))

def add_item():
    particulars = particulars_entry.get()
    quantity = int(quantity_entry.get())
    unit_price = float(price_entry.get())
    
    tree_items = tree.get_children()
    invoice_list = []
    
    for i in tree_items:
        item = tree.item(i)['values']
        invoice_list.append(item[0].lower())
       
    if(particulars == ""):
        messagebox.showerror(parent=invoice_tk, title="Invalid Product Name", message="Please enter a valid product name")
    
    elif(quantity < 1):
        messagebox.showerror(parent=invoice_tk, title="Invalid Quantity", message="Please enter a valid quantity count")
    
    elif(unit_price <= 0.0):
        messagebox.showerror(parent=invoice_tk, title="Invalid Unit Price", message="Please enter a valid unit price")
    
    else:
        if(len(tree.selection())<1):
            if(particulars.lower() in invoice_list):
                messagebox.showerror(parent=invoice_tk, title="Product Already Exists", message="Entered product already exists in the invoice")
            else:
                particulars = particulars_entry.get()
                quantity = int(quantity_entry.get())
                unit_price = float(price_entry.get())
                line_total = quantity*unit_price
                invoice_item = [particulars, quantity, unit_price, line_total]
                tree.insert('',0, values=invoice_item)
                clear_item()
            
        else:
            selected_item = tree.selection()[0]  
            tree.delete(selected_item)
                
            particulars = particulars_entry.get()
            quantity = int(quantity_entry.get())
            unit_price = float(price_entry.get())
            line_total = quantity*unit_price
            invoice_item = [particulars, quantity, unit_price, line_total]
            tree.insert('',0, values=invoice_item)
            clear_item()
        
        amount_calculator()
    
def generate_invoice():
    try:
        if(name_entry.get() == ""):
            messagebox.showerror(parent=invoice_tk, title="Invalid Name", message="Please enter a valid name")
    
        elif(address_entry.get() == ""):
            messagebox.showerror(parent=invoice_tk, title="Invalid Address", message="Please enter a valid address")
    
        elif(phone_number_entry.get() == "" or len(phone_number_entry.get())!=10):
            messagebox.showerror(parent=invoice_tk, title="Invalid Phone Number", message="Please enter a valid phone number")
    
        elif(gst_entry.get() == "" or len(gst_entry.get())!=15):
            messagebox.showerror(parent=invoice_tk, title="Invalid GST Number", message="Please enter a valid GST number")
    
        elif(len(tree.get_children()) == 0):
            messagebox.showerror(parent=invoice_tk, title="Empty Invoice", message="Please add items to the invoice")
    
        else:
            doc = DocxTemplate("invoice_template.docx")
            name = name_entry.get()
            address = address_entry.get()
            phone = phone_number_entry.get()
            gst = gst_entry.get()
            inno = "IN-2425-{number:06}".format(number=invoice_no)
            current_date = date.today()
            t = time.localtime()
            current_time = time.strftime("%H:%M:%S", t)
            
            tree_items = tree.get_children()
            invoice_list = []
            
            for i in tree_items:
                item = tree.item(i)['values']
                invoice_list.append(item)
            
            subtotal = sum(float(item[3]) for item in invoice_list)
            
            cgst = subtotal*0.025
            sgst = subtotal*0.025
            gtotal = subtotal+cgst+sgst
            
            amount_text = (number_to_words_inr(gtotal)).capitalize()
        
            doc.render({"name":name,"address":address,"phone":phone,"gst":gst,
                "invoiceno":inno,"invoicedate":current_date,"invoicetime":current_time,
                "invoice_list": invoice_list,"subtotal":subtotal,"cgst":cgst,
                "sgst":sgst,"gtotal":gtotal, "text_amount":amount_text})
    
            particulars_data = ""
            quantity_data = ""
            price_data = ""
            amount_data = ""
    
            for i in range(len(invoice_list)):
                particulars_data += invoice_list[i][0]
                quantity_data += str(invoice_list[i][1])
                price_data += str(invoice_list[i][2])
                amount_data += str(invoice_list[i][3])
        
                if(i<len(invoice_list)-1):
                    particulars_data += ","
                    quantity_data += ","
                    price_data += ","
                    amount_data += ","
    
            insert = ("INSERT INTO invoice VALUES('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}')").format(inno,current_date,current_time,name,address,phone,gst,particulars_data,quantity_data,price_data,amount_data,str(subtotal),str(cgst),str(sgst),str(gtotal))
            cursor.execute(insert)
            mycon.commit()
    
            insert = ("INSERT INTO accounts VALUES('{particulars}',{amount},'{date}')").format(date=current_date,particulars=inno,amount=gtotal)
            cursor.execute(insert)
            mycon.commit()     
           
            doc_name = "IN-2425-{number:06}".format(number=invoice_no)+".docx"
            doc.save('F:/ABC Project/Consultancy Project/Project/Invoices/'+doc_name)
    
            answer = messagebox.askyesno(parent=invoice_tk,title="Print invoice", message="Do you want to print the invoice?")
    
            if(answer == True):
                print_document()
        
            messagebox.showinfo(parent=invoice_tk,title="Invoice Saved", message="Invoice Saved Successfully")
    
    except:
        messagebox.showerror(parent=invoice_tk,title="Duplicate Invoice", message="Invoice already exists in the database. Please create a new invoice.")

def new_invoice():
    answer = messagebox.askyesno(parent=invoice_tk,title="New Invoice", message="Do you want to create a new invoice?")
    
    if(answer == True):
        name_entry.delete(0,END)
        address_entry.delete(0,END)
        phone_number_entry.delete(0,END)
        gst_entry.delete(0,END)
    
        particulars_entry.delete(0,END)
        quantity_entry.delete(0,END)
        quantity_entry.insert(0, "0")
        price_entry.delete(0,END)
        price_entry.insert(0, "0.0")
    
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
        
        show = ("SELECT COUNT(*) FROM invoice")
        cursor.execute(show)
        invoice_no=cursor.fetchall()
        invoice_no=int(invoice_no[0][0])+1
        
        Label(invoice_tk,text="Invoice Number : IN-2425-{number:06}".format(number=invoice_no),font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1250,y=20)

def print_document():
    file_path = "F:/ABC Project/Consultancy Project/Project/Invoices/"+"IN-2425-{number:06}".format(number=invoice_no)+".docx"
    printer_path = 'Microsoft Print to PDF'
    file_handle = open(file_path, 'rb')
    
    printer_handle = win32print.OpenPrinter(printer_path)
    JobInfo = win32print.StartDocPrinter(printer_handle, 1, (file_path, None, "RAW"))
    win32print.StartPagePrinter(printer_handle)
    win32print.WritePrinter(printer_handle, file_handle.read())
    win32print.EndPagePrinter(printer_handle)
            
def edit_invoice():
    os.system('python invoice_edit.py')

def view_invoice():
    os.system('python invoice_view.py')
    
Label(invoice_tk,text="AJRA TEX - KARUR",font=("Arial", 20, "bold"),bg="white",fg="#1A374D").place(x=650,y=20)
Label(invoice_tk,text="Invoice Number : IN-2425-{number:06}".format(number=invoice_no),font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1250,y=20)
'''------- Buyer details -------'''

name_label = Label(invoice_tk, text="Name",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=260,y=100)
name_entry = Entry(invoice_tk,font=("Arial", 12),bd=3)
name_entry.place(x=190,y=130)

address_label = Label(invoice_tk, text="Address",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=560,y=100)
address_entry = Entry(invoice_tk,font=("Arial", 12),bd=3)
address_entry.place(x=500,y=130)

phone_number_label = Label(invoice_tk, text="Phone number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=850,y=100)
phone_number_entry = Entry(invoice_tk,font=("Arial", 12),bd=3)
phone_number_entry.place(x=810,y=130)

gst_label = Label(invoice_tk, text="GST number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1160,y=100)
gst_entry = Entry(invoice_tk,font=("Arial", 12),bd=3)
gst_entry.place(x=1110,y=130)

'''------- Product details -------'''

particulars_label = Label(invoice_tk, text="Particulars",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=240,y=180)
particulars_entry = Entry(invoice_tk,font=("Arial", 12),bd=3)
particulars_entry.place(x=190,y=210)

quantity_label = Label(invoice_tk, text="Quantity",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=560,y=180)
quantity_entry = Spinbox(invoice_tk, from_=0,to=100,font=("Arial", 12),bd=3)
quantity_entry.place(x=500,y=210)

price_label = Label(invoice_tk, text="Unit Price",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=870,y=180)
price_entry = Spinbox(invoice_tk,from_=0.0, to=5000, increment=0.5,font=("Arial", 12),bd=3)
price_entry.place(x=810,y=210)

add_item_button = Button(invoice_tk, text="Add item",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",command=add_item)
add_item_button.place(x=1110,y=210)

columns = ('particulars', 'quantity', 'unit_price', 'amount')
tree = ttk.Treeview(invoice_tk, columns=columns, show="headings")
tree.heading('particulars', text='Particulars')
tree.heading('quantity', text='Quantity')
tree.heading('unit_price', text='Unit Price')
tree.heading('amount', text="Amount")
tree.place(x=190,y=300)

stotal_label = Label(invoice_tk, text="Sub Total : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=300)
stotal_entry = Entry(invoice_tk,font=("Arial", 12),bd=3)
stotal_entry.place(x=1220,y=300)
stotal_entry.insert(0,"0.0")

cgst_label = Label(invoice_tk, text="CGST @2.5% : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=350)
cgst_entry = Entry(invoice_tk,font=("Arial", 12),bd=3)
cgst_entry.place(x=1220,y=350)
cgst_entry.insert(0,"0.0")

sgst_label = Label(invoice_tk, text="SGST @2.5% : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=400)
sgst_entry = Entry(invoice_tk,font=("Arial", 12),bd=3)
sgst_entry.place(x=1220,y=400)
sgst_entry.insert(0,"0.0")

gtotal_label = Label(invoice_tk, text="Grand Total : ",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1100,y=500)
gtotal_entry = Entry(invoice_tk,font=("Arial", 12, "bold"),bd=3)
gtotal_entry.place(x=1220,y=500)
gtotal_entry.insert(0,"0.0")

save_invoice_button = Button(invoice_tk, text="Generate Invoice",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=generate_invoice)
save_invoice_button.place(x=650,y=580)

new_invoice_button = Button(invoice_tk, text="New Invoice",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=new_invoice)
new_invoice_button.place(x=650,y=640)

edit_invoice_button = Button(invoice_tk, text="Edit Invoice",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=edit_invoice)
edit_invoice_button.place(x=650,y=700)

view_invoice_button = Button(invoice_tk, text="View Invoice",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=view_invoice)
view_invoice_button.place(x=650,y=760)

edit_btn = ttk.Button(invoice_tk,text="Edit", command=edit_selection)
edit_btn.place(x=915,y=530)

delete_btn = ttk.Button(invoice_tk,text="Delete", command=delete_selection)
delete_btn.place(x=830,y=530)

invoice_tk.mainloop()
