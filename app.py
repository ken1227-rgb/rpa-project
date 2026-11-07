# problem1
while True:
    directory = single_input("Please input the directory where the file is located.", mask=False, width=400, height=200)
    if directory.endswith('/'):
        pass
    else:
        directory=directory+'/'
    filenames = os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames:
        break
    else:
        show_message("Error", 'Not a valid directory')

#problem2
df1=pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')
df2=pd.read_excel(directory+'ProductData.xlsx',sheet_name='Prices')
product_descriptions = {}
product_prices={}
for index,row in df1.iterrows():
    number=row['Unit number']
    description=row['Description']
    product_descriptions[number]=description
for index,row in df2.iterrows():
    number=row['Unit number']
    price=row['Price per unit']
    product_prices[number]=price
#problem3
def get_order_data(invoice_num,filename,directory):
    order_data={}
    workbook=openpyxl.load_workbook(directory+filename)
    selected_sheet=workbook['OrderForm']
    email=selected_sheet['C13'].value
    name=selected_sheet['C14'].value
    date=selected_sheet['C9'].value.strftime('%m/%d/20%y')
    address=selected_sheet['C16'].value
    city=selected_sheet['C17'].value
    state=selected_sheet['C18'].value
    zipcode=selected_sheet['C19'].value
    companyname=selected_sheet['C15'].value
    cell_number=24
    Total=0
    productnumber=0
    while True:
        unit_number=selected_sheet[f'B{cell_number}'].value
        if unit_number is None:
            break
        else:
            quantity=selected_sheet[f'C{cell_number}'].value
            product_price=product_prices[unit_number]
            subtotal=quantity*product_price
            description=product_descriptions[unit_number]
            Total=Total+subtotal
            cell_number=cell_number+1
            order_data[f'Price{productnumber}']=f'${product_price:,.2f}'
            order_data[f'Quantity{productnumber}']=f'{quantity:,.2f}'
            order_data[f'Amount{productnumber}']=f'${subtotal:,.2f}'
            order_data[f'Description{productnumber}']=f'{description}'
            productnumber=productnumber+1
   
    Total=f'${Total:,.2f}'
    order_data['InvoiceNum']=invoice_num
    order_data['Date']=date
    order_data['Total']=Total
    order_data['CustomerAddress1']=f'{address}'
    order_data['CustomerAddress2']=f'{city},{state},{zipcode}'
    order_data['CustomerName']=name
    order_data['CustomerCompany']=companyname
    
    return order_data, name, email,date
#problem4
def create_invoices(invoice_num,order_data,directory):
    create_pdf(directory+'MasterInvoice.pdf',f'{directory}Invoice{invoice_num}.pdf',order_data)
#problem 5
def send_email(invoice_num,name,email,date,directory):
    # Enter your gmail user name and password

    gmail_user = "ken1227sei.nono@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    

    body = f""" Dear {name},
        
    Attached is your invoice dated {str(date)}.
    
    Thank you for your business.

    PythonTech Co.
    10Python Drive
    Pleasant Grove UT 84602
    Phone: (801)555-2372
    Email: sales@pythontechco.com"""
        
    

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    # path_to_file = os.getcwd()+'\\Invoice Template.xlsx'
    path_to_file = directory+f'Invoice{invoice_num}.pdf' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

    return
#problem 6
filenames = os.listdir(directory)
filenames=sorted(filenames)
invoice_num=8515
for item in filenames:
    if 'Order' in item and 'xlsx' in item:
        order_data,name,email,date=get_order_data(invoice_num,item,directory)
        create_invoices(invoice_num,order_data,directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num=invoice_num+1