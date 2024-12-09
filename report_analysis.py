from tkinter import *
from tkinter import ttk
import mysql.connector as sql
from tkinter import Canvas

ra_tk=Tk()
ra_tk.geometry("1920x1080")
ra_tk.config(bg="#F5F5F5")
ra_tk.title("Report and Analysis - Ajra Tex")
ra_tk.configure(bg='white')

mycon=sql.connect(host="localhost",user="root",passwd="gokul123",database="ajra")
cursor=mycon.cursor()
           
months = ['January', 'February', 'March', 'April', 'May', 'June', 'July','August', 'September', 'October', 'November', 'December']
sales_list = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
purchase_list = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

show = ("SELECT SUM(amount), MONTH(date) FROM accounts WHERE amount > 0 GROUP BY MONTH(date)")
cursor.execute(show)
sales=cursor.fetchall()

show = ("SELECT SUM(amount), MONTH(date) FROM accounts WHERE amount < 0 GROUP BY MONTH(date)")
cursor.execute(show)
purchase=cursor.fetchall()


for i in range(len(sales)):
    sales_list[sales[i][1]-1] = sales[i][0]
for i in range(len(purchase)):
    purchase_list[purchase[i][1]-1] = (purchase[i][0])*(-1)
   
data = []

for i in range(12):
    tup = (months[i], sales_list[i], purchase_list[i])
    data.append(tup)

Label(ra_tk,text="AJRA TEX - KARUR",font=("Arial", 20, "bold"),bg="white",fg="#1A374D").place(x=650,y=20)

'''---------- Monthly Income Report ----------'''

Label(ra_tk,text="Monthly Income Report",font=("Arial", 15, "bold"),bg="white",fg="#1A374D").place(x=190,y=100)

columns = ('months', 'income')
tree = ttk.Treeview(ra_tk, columns=columns, show="headings")
tree.heading('months', text='Months')
tree.heading('income', text='Income')
tree.place(x=100,y=150) 
    
for i in range(12):
    tree.insert('',i, values=[data[i][0],data[i][1]])
    
'''---------- Monthly Expence Report ----------'''

Label(ra_tk,text="Monthly Expence Report",font=("Arial", 15, "bold"),bg="white",fg="#1A374D").place(x=190,y=450)
columns = ('months', 'expences')
tree = ttk.Treeview(ra_tk, columns=columns, show="headings")
tree.heading('months', text='Months')
tree.heading('expences', text='Expences')
tree.place(x=100,y=500)           
    
for i in range(12):
    tree.insert('',i, values=[data[i][0],data[i][2]])
    



'''---------- Graph ----------'''
CANVAS_WIDTH = 750
CANVAS_HEIGHT = 550
BAR_WIDTH = 40
BAR_SPACING = 20
CHART_TOP_MARGIN = 30

def draw_bar_chart(canvas, data):
    max_value = max(max(item[1], item[2]) for item in data)

    scale = (CANVAS_HEIGHT - CHART_TOP_MARGIN) / max_value
    x = BAR_SPACING

    for category, value1, value2 in data:
        bar_height1 = value1 * scale
        bar_height2 = value2 * scale
        canvas.create_rectangle(x, CANVAS_HEIGHT - bar_height1, x + BAR_WIDTH, CANVAS_HEIGHT,fill="#21a1f1", outline="white")
        canvas.create_rectangle(x, CANVAS_HEIGHT - bar_height2, x + BAR_WIDTH, CANVAS_HEIGHT,fill="#d90d1b", outline="white")
        canvas.create_text(x + BAR_WIDTH // 2, CANVAS_HEIGHT - bar_height1 - 10, text=category)
        
        x += BAR_WIDTH + BAR_SPACING
            
canvas = Canvas(ra_tk, width=CANVAS_WIDTH, height=CANVAS_HEIGHT, bg="white")
canvas.place(x=650,y=150)
draw_bar_chart(canvas, data)

Label(ra_tk,text="    ",bg="#21a1f1",fg="#21a1f1").place(x=700,y=720)
Label(ra_tk,text="Income",font=("Arial", 12),bg="white",fg="black").place(x=730,y=718)

Label(ra_tk,text="    ",bg="#d90d1b",fg="#d90d1b").place(x=900,y=720)
Label(ra_tk,text="Expences",font=("Arial", 12),bg="white",fg="black").place(x=930,y=718)

ra_tk.mainloop()