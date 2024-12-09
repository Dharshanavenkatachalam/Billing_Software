from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
from datetime import date
import time
from tkinter import messagebox
import mysql.connector as sql
import win32print

iedit_tk=Tk()
iedit_tk.geometry("1920x1080")
iedit_tk.config(bg="#F5F5F5")
iedit_tk.title("Edit Invoice - Ajra Tex")
iedit_tk.configure(bg='white')

mycon=sql.connect(host="localhost",user="root",passwd="gokul123",database="ajra")
cursor=mycon.cursor()

def clear_item():
    particulars_entry.delete(0,END)
    quantity_entry.delete(0,END)
    quantity_entry.insert(0, "0")
    price_entry.delete(0,END)
    price_entry.insert(0, "0.0")

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

def add_item():
    tree_items = tree.get_children()
    invoice_list = []
    
    for i in tree_items:
        item = tree.item(i)['values']
        invoice_list.append(item[0].lower())
        
    if(particulars_entry.get()==""):
        messagebox.showerror(parent=iedit_tk, title="Invalid particulars", message="Please enter a valid particulars")
        return
    
    elif(float(quantity_entry.get())<=0 or quantity_entry.get() == ""):
        messagebox.showerror(parent=iedit_tk, title="Invalid quantity", message="Please enter a valid quantity")
        return
    
    elif(float(price_entry.get())<=0 or price_entry.get() == ""):
        messagebox.showerror(parent=iedit_tk, title="Invalid price", message="Please enter a valid price")
        return
    
    else:
        if(len(tree.selection())<1):
            if((particulars_entry.get()).lower() in invoice_list):
                messagebox.showerror(parent=iedit_tk, title="Product Already Exists", message="Entered product already exists in the invoice")
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
        
def edit_invoice():
    try:
        clear_entry()
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
         messagebox.showerror(tilte="Invalid Invoice Number", message="Please enter a valid invoice number")
        
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

def generate_invoice():
    if(name_entry.get() == ""):
        messagebox.showerror(parent=iedit_tk, title="Invalid Name", message="Please enter a valid name")
        return
    
    elif(address_entry.get() == ""):
        messagebox.showerror(parent=iedit_tk, title="Invalid Address", message="Please enter a valid address")
        return
    
    elif(phone_number_entry.get() == "" or len(phone_number_entry.get())!=10):
        messagebox.showerror(parent=iedit_tk, title="Invalid Phone Number", message="Please enter a valid phone number")
        return
    
    elif(gst_entry.get() == "" or len(gst_entry.get())!=15):
        messagebox.showerror(parent=iedit_tk, title="Invalid GST Number", message="Please enter a valid GST number")
        return 
    
    elif(len(tree.get_children()) == 0):
        messagebox.showerror(parent=iedit_tk, title="Empty Invoice", message="Please add items to the invoice")
        return
    
    else:
        doc = DocxTemplate("invoice_template.docx")
        name = name_entry.get()
        address = address_entry.get()
        phone = phone_number_entry.get()
        gst = gst_entry.get()
        inno = invoice_number_entry.get()
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
        
        doc.render({"name":name,
                "address":address,
                "phone":phone,
                "gst":gst,
                "invoiceno":inno,
                "invoicedate":current_date,
                "invoicetime":current_time,
                "invoice_list": invoice_list,
                "subtotal":subtotal,
                "cgst":cgst,
                "sgst":sgst,
                "gtotal":gtotal,"text_amount":amount_text})
    
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
    
        update_invoice = ("UPDATE invoice SET invoice_date='{}',invoice_time='{}',name='{}',address='{}',phone_no='{}',gst_no='{}',particulars='{}',quantity='{}',price='{}',amount='{}',sub_total='{}',cgst='{}',sgst='{}',grand_total='{}' WHERE invoice_no='{}'").format(current_date,current_time,name,address,phone,gst,particulars_data,quantity_data,price_data,amount_data,str(subtotal),str(cgst),str(sgst),str(gtotal),inno)
        cursor.execute(update_invoice)
        mycon.commit()
    
        update_accounts = ("UPDATE accounts SET amount = '{}',date = '{}' WHERE description = '{}'").format(gtotal,current_date,inno)
        cursor.execute(update_accounts)
        mycon.commit()
    
        doc_name = inno+".docx"
        doc.save('F:/ABC Project/Consultancy Project/Project/Invoices/'+doc_name)
        answer = messagebox.askyesno("Print invoice", "Do you want to print the invoice?")
    
        if(answer == True):
            print_document()
                   
        messagebox.showinfo(parent=iedit_tk, title="Invoice Edited", message="Invoice Edited Successfully")

def print_document():
    file_path = "F:/ABC Project/Consultancy Project/Project/Invoices/"+invoice_number_entry.get()+".docx"
    printer_path = 'Microsoft Print to PDF'
    file_handle = open(file_path, 'rb')
    
    printer_handle = win32print.OpenPrinter(printer_path)
    JobInfo = win32print.StartDocPrinter(printer_handle, 1, (file_path, None, "RAW"))
    win32print.StartPagePrinter(printer_handle)
    win32print.WritePrinter(printer_handle, file_handle.read())
    win32print.EndPagePrinter(printer_handle)

def new_invoice():
    answer = messagebox.askyesno(parent=iedit_tk, title="New Invoice", message="Do you want to create a new invoice?")
    
    if(answer == True):
        name_entry.delete(0,END)
        address_entry.delete(0,END)
        phone_number_entry.delete(0,END)
        gst_entry.delete(0,END)
    
        invoice_number_entry.delete(0,END)
    
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
        
def clear_entry():
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

def delete_selection():
    if(len(tree.selection())<1):
        messagebox.showerror(parent=iedit_tk, title="Selection Error", message="Please select an item to edit")
        return
    
    else:
        answer = messagebox.askyesno(parent=iedit_tk, title="Delete Item", message="Are you sure to delete the selected item?")
        
        if(answer == True):
            selected_item = tree.selection()[0]
            tree.delete(selected_item)
        
        amount_calculator()
    
Label(iedit_tk,text="AJRA TEX - KARUR",font=("Arial", 20, "bold"),bg="white",fg="#1A374D").place(x=650,y=20)
Label(iedit_tk,text="Invoice Number : ",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=580,y=100)

invoice_number_entry = Entry(iedit_tk,font=("Arial", 12, "bold"),bd=3)
invoice_number_entry.place(x=730,y=100)

edit_button = Button(iedit_tk, text="Submit",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",command=edit_invoice)
edit_button.place(x=930,y=97)

name_label = Label(iedit_tk, text="Name",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=260,y=150)
name_entry = Entry(iedit_tk,font=("Arial", 12),bd=3)
name_entry.place(x=190,y=180)

address_label = Label(iedit_tk, text="Address",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=560,y=150)
address_entry = Entry(iedit_tk,font=("Arial", 12),bd=3)
address_entry.place(x=500,y=180)

phone_number_label = Label(iedit_tk, text="Phone number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=850,y=150)
phone_number_entry = Entry(iedit_tk,font=("Arial", 12),bd=3)
phone_number_entry.place(x=810,y=180)

gst_label = Label(iedit_tk, text="GST number",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1160,y=150)
gst_entry = Entry(iedit_tk,font=("Arial", 12),bd=3)
gst_entry.place(x=1110,y=180)

'''------- Product details -------'''

particulars_label = Label(iedit_tk, text="Particulars",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=240,y=230)
particulars_entry = Entry(iedit_tk,font=("Arial", 12),bd=3)
particulars_entry.place(x=190,y=260)

quantity_label = Label(iedit_tk, text="Quantity",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=560,y=230)
quantity_entry = Spinbox(iedit_tk, from_=0,to=100,font=("Arial", 12),bd=3)
quantity_entry.place(x=500,y=260)

price_label = Label(iedit_tk, text="Unit Price",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=870,y=230)
price_entry = Spinbox(iedit_tk,from_=0.0, to=5000, increment=0.5,font=("Arial", 12),bd=3)
price_entry.place(x=810,y=260)

add_item_button = Button(iedit_tk, text="Add item",font=("Arial", 10,"bold"),bg="#1A374D",fg="#F5F5F5",command=add_item)
add_item_button.place(x=1110,y=255)

columns = ('particulars', 'quantity', 'unit_price', 'amount')
tree = ttk.Treeview(iedit_tk, columns=columns, show="headings")
tree.heading('particulars', text='Particulars')
tree.heading('quantity', text='Quantity')
tree.heading('unit_price', text='Unit Price')
tree.heading('amount', text="Amount")
tree.place(x=190,y=350)

edit_btn = ttk.Button(iedit_tk,text="Edit", command=edit_selection)
edit_btn.place(x=915,y=580)

delete_btn = ttk.Button(iedit_tk,text="Delete", command=delete_selection)
delete_btn.place(x=830,y=580)

stotal_label = Label(iedit_tk, text="Sub Total : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=350)
stotal_entry = Entry(iedit_tk,font=("Arial", 12),bd=3)
stotal_entry.place(x=1220,y=350)
stotal_entry.insert(0,"0.0")

cgst_label = Label(iedit_tk, text="CGST @2.5% : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=400)
cgst_entry = Entry(iedit_tk,font=("Arial", 12),bd=3)
cgst_entry.place(x=1220,y=400)
cgst_entry.insert(0,"0.0")

sgst_label = Label(iedit_tk, text="SGST @2.5% : ",font=("Arial", 12),bg="white",fg="#1A374D").place(x=1100,y=450)
sgst_entry = Entry(iedit_tk,font=("Arial", 12),bd=3)
sgst_entry.place(x=1220,y=450)
sgst_entry.insert(0,"0.0")

gtotal_label = Label(iedit_tk, text="Grand Total : ",font=("Arial", 12,"bold"),bg="white",fg="#1A374D").place(x=1100,y=500)
gtotal_entry = Entry(iedit_tk,font=("Arial", 12, "bold"),bd=3)
gtotal_entry.place(x=1220,y=500)
gtotal_entry.insert(0,"0.0")

save_invoice_button = Button(iedit_tk, text="Generate Invoice",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=generate_invoice)
save_invoice_button.place(x=650,y=640)

new_invoice_button = Button(iedit_tk, text="New Invoice",font=("Arial",10,"bold"),bg="#1A374D",fg="#F5F5F5",width=30,height=2,command=new_invoice)
new_invoice_button.place(x=650,y=700)

iedit_tk.mainloop()
