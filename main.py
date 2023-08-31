import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import socket

def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")

invoice_list = []


def add_item():
    qty = int(qty_spinbox.get())
    part_number = part_number_entry.get()
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    line_total = qty * price
    invoice_item = [qty, part_number, desc, price, line_total]
    print(invoice_item)
    tree.insert("", 0, values=invoice_item)
    clear_item()
    
    invoice_list.append(invoice_item)
    update_total()

def update_total():
    subtotal = sum(item[4] for item in invoice_list)
    discount = float(discount_spinbox.get())
    discounted_subtotal = subtotal - discount
    total_value.config(text=f"{discounted_subtotal:.2f}")

def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    total_value.config(text="0.00")
    discount_spinbox.delete(0, tkinter.END)
    discount_spinbox.insert(0, '0.0')
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()

# def has_internet():
#     try:
#         # Connect to a well-known host using a well-known port
#         # This will fail if there's no internet connection
#         socket.create_connection(("www.google.com", 80))
#         return True
#     except OSError:
#         return False

def generate_invoice():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_file_path = os.path.join(script_dir, "invoice_template.docx")
    invoices_folder = os.path.join(script_dir, "invoices")
    if not os.path.exists(invoices_folder):
        os.mkdir(invoices_folder)
    
    name = first_name_entry.get() + " " + last_name_entry.get()
    phone = phone_entry.get()
    
    # Calculate the subtotal based on line totals (quantity * price) of each item
    subtotal = sum(item[4] for item in invoice_list)  # Index 4 is the line total
    
    salestax = '23AGAPC5563E1ZZ'
    discount = float(discount_spinbox.get())
    discounted_subtotal = subtotal - discount
    total = discounted_subtotal
    
    doc = DocxTemplate(template_file_path)
    doc.render({"name": name,
                "dis": discount,
                "phone": phone,
                "invoice_list": invoice_list,
                "subtotal": subtotal,
                "salestax": salestax,
                "total": total})
    
    doc_name = os.path.join(invoices_folder, f"{name}_{datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')}.docx")
    doc.save(doc_name)
    os.startfile(doc_name, 'print')
    
    send_email(name, doc_name)
    
    messagebox.showinfo("Invoice Complete", "Invoice Complete")
    
    new_invoice()


def send_email(customer_name, invoice_file):
    from_email = "mohitchhabaria160@gmail.com"  # Replace with your email
    to_email = "k.d.tvsmandideep@gmail.com"  # Replace with customer's email
    cc_email = "mohitchhabaria899@gmail.com"  # Replace with CC email address
    subject = "Invoice for Your Purchase"
    body = f"Dear {customer_name},\n\nPlease find attached the invoice for your recent purchase.\n\nBest regards,\nYour K.D. Auto Mobile"
    
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Cc'] = cc_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    with open(invoice_file, "rb") as f:
        attachment = MIMEApplication(f.read(), _subtype="docx")
        attachment.add_header('Content-Disposition', f'attachment; filename={invoice_file}')
        msg.attach(attachment)
    
    smtp_server = 'smtp.gmail.com'  # Replace with your SMTP server
    smtp_port = 587  # Replace with your SMTP port
    
    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(from_email, "jpmkxilbikaozavm")  # Replace with your email password
    smtp.sendmail(from_email, [to_email, cc_email], msg.as_string())
    smtp.quit()


def apply_discount():
    update_total()

# Create the main window
window = tkinter.Tk()
window.title("K.D.TVS")
window.configure(bg="grey")

frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10)

first_name_label = tkinter.Label(frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(frame)
last_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

phone_label = tkinter.Label(frame, text="Phone")
phone_label.grid(row=0, column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1, column=2)

qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row=2, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

desc_label = tkinter.Label(frame, text="Part name")
desc_label.grid(row=2, column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)

price_label = tkinter.Label(frame, text="Price")
price_label.grid(row=2, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500000, increment=1)
price_spinbox.grid(row=3, column=2)

discount_label = tkinter.Label(frame, text="Discount")
discount_label.grid(row=4, column=3)

discount_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500000, increment=1)
discount_spinbox.grid(row=5, column=3)

apply_discount_button = tkinter.Button(frame, text="Apply Discount", command=apply_discount, bg='skyblue')
apply_discount_button.grid(row=5, column=4)

add_item_button = tkinter.Button(frame, text="Add item", command=add_item, bg='green')
add_item_button.grid(row=6, column=2, pady=5)

columns = ('Quantity', 'part number', 'Part name', 'price', 'line total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('Quantity', text='Qty')
tree.heading('Part name', text='Part Name')
tree.heading('price', text='Price')
tree.heading('line total', text="Total")

tree.grid(row=7, column=0, columnspan=4, padx=20, pady=10)

total_label = tkinter.Label(frame, text="Total:")
total_label.grid(row=8, column=0, sticky='e', padx=10)

total_value = tkinter.Label(frame, text="0.0")
total_value.grid(row=8, column=1, columnspan=2, sticky='w')

save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice, bg='green')
save_invoice_button.grid(row=9, column=0, columnspan=2, sticky="news", padx=20, pady=5)

new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice, bg='green')
new_invoice_button.grid(row=9, column=2, columnspan=2, sticky="news", padx=20, pady=5)

part_number_label = tkinter.Label(frame, text="Part Number")
part_number_label.grid(row=2, column=3)
part_number_entry = tkinter.Entry(frame)
part_number_entry.grid(row=3, column=3)

def delete_item():
    selected_item = tree.selection()
    if selected_item:
        item_id = selected_item[0]
        tree.delete(item_id)
        index = int(item_id[1:]) - 1
        print(index)
        if 0 <= index < len(invoice_list):
            invoice_list.pop(index)
            print(invoice_list)
            update_total()

tree.column("#4", stretch=tkinter.NO)
tree.heading('#4', text='Delete')
for item in tree.get_children():
    tree.insert(item, 'end', values=["", "", "", "", "Delete"])
for col in columns:
    tree.column(col, anchor='center')
    tree.heading(col, text=col, anchor='center')

delete_button = tkinter.Button(frame, text="Delete", command=delete_item, bg='red')
delete_button.grid(row=5, column=2, padx=5)

window.mainloop()
