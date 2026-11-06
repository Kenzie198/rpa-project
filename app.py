while True: 
    directory = single_input("Please input the directory where the 'ProductData.xlsx', 'MasterInvoice.pdf, and 'Order' files are located.", mask=False, width=400, height=200)
    directory = directory.strip('"')

    if directory[-1] != '\\':
        directory = directory + '\\'  
   
    if ('ProductData.xlsx') and ('MasterInvoice.pdf') and ('Order1.xlsx') and ('Order2.xlsx') and ('Order3.xlsx') and ('Order4.xlsx') in os.listdir(directory):
        break
    else:
        show_message("Error", "The directory you entered does not contain the 'ProductData.xlsx', 'MasterInvoice.pdf, and 'Order' files. Please try again.")

directory

product_descriptions = {}
product_prices = {}
descriptions = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Descriptions')
for index,row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

    
prices = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Prices')
for index,row in prices.iterrows():
    product_prices[row['Unit number']] = str(row['Price per unit'])

product_prices

def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = selected_sheet['C9'].value.strftime('%m/%d/%Y')

    if selected_sheet['B24'].value is not None:
        order_data['Unit Number0'] = selected_sheet['B24'].value
        order_data['Price0'] = float((product_prices[selected_sheet['B24'].value]))
        order_data['Quantity0'] = (selected_sheet['C24'].value)
        order_data['Amount0'] = (order_data['Price0']) * (order_data['Quantity0'])
        order_data['Description0'] = str(product_descriptions[selected_sheet['B24'].value])
        order_data['Total'] = (order_data['Price0'] * order_data['Quantity0'])

    if selected_sheet['B25'].value is not None:
        order_data['Unit Number1'] = selected_sheet['B25'].value
        order_data['Price1'] = float((product_prices[selected_sheet['B25'].value]))
        order_data['Quantity1'] = (selected_sheet['C25'].value)
        order_data['Amount1'] = (order_data['Price1']) * (order_data['Quantity1'])
        order_data['Description1'] = str(product_descriptions[selected_sheet['B25'].value])
        order_data['Total'] = (order_data['Price0'] * order_data['Quantity0']) + (order_data['Price1'] * order_data['Quantity1'])

        
    if selected_sheet['B26'].value is not None:
        order_data['Unit Number2'] = selected_sheet['B26'].value
        order_data['Price2'] = float((product_prices[selected_sheet['B26'].value]))
        order_data['Quantity2'] = (selected_sheet['C26'].value)
        order_data['Amount2'] = (order_data['Price2']) * (order_data['Quantity2'])
        order_data['Description2'] = str(product_descriptions[selected_sheet['B26'].value])
        order_data['Total'] = (order_data['Price0'] * order_data['Quantity0']) + (order_data['Price1'] * order_data['Quantity1']) + (order_data['Price2'] * order_data['Quantity2'])
        

    if selected_sheet['B24'].value is not None:
        order_data['Price0'] = f'${order_data['Price0']:,.2f}'
        order_data['Amount0'] = f'${order_data['Amount0']:,.2f}'    
        order_data['Quantity0'] = f'{(order_data['Quantity0'])}'

    if selected_sheet['B25'].value is not None:
        order_data['Price1'] = f'${order_data['Price1']:,.2f}'
        order_data['Amount1'] = f'${order_data['Amount1']:,.2f}'
        order_data['Quantity1'] = f'{(order_data['Quantity1'])}'

    if selected_sheet['B26'].value is not None:
        order_data['Price2'] = f'${order_data['Price2']:,.2f}'
        order_data['Amount2'] = f'${order_data['Amount2']:,.2f}'
        order_data['Quantity2'] = f'{(order_data['Quantity2'])}'
    order_data['Total'] = f'${(order_data['Total']):,.2f}'
            
    # if selected_sheet['B25'].value is None:
    #     order_data['Total'] = (order_data['Price0'] * order_data['Quantity0'])
    # elif selected_sheet['B26'].value is None:
    #     order_data['Total'] = (order_data['Price0'] * order_data['Quantity0']) + (order_data['Price1'] * order_data['Quantity1'])
    # else:
    #     order_data['Total'] = (order_data['Price0'] * order_data['Quantity0']) + (order_data['Price1'] * order_data['Quantity1']) + (order_data['Price2'] * order_data['Quantity2'])
    order_data['CustomerAddress1'] = (selected_sheet['C16'].value)
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = str(selected_sheet['C17'].value) + ', ' + str(selected_sheet['C18'].value) + ' ' + str(selected_sheet['C19'].value)

    # order_data['Quantity0'] = f'{(order_data['Quantity0'])}'
    # order_data['Quantity1'] = f'{(order_data['Quantity1'])}'
    # order_data['Quantity2'] = f'{(order_data['Quantity2'])}'
    # order_data['Description0'] = f'{(order_data['Description0'])}'
    # order_data['Description1'] = f'{(order_data['Description1'])}'
    # order_data['Description2'] = f'{(order_data['Description2'])}'

    name = order_data['CustomerName']

    email = selected_sheet['C13'].value

    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    
    return order_data, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
order_data, name, email, date

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)

new_invoice = create_invoice('8515',order_data,directory)
print(new_invoice)

gmail_password_path = select_file('Please select the gmail password file')

gmail_password = open(gmail_password_path).read()   


def send_email (invoice_num, name, email, date, directory):

    gmail_user = "YOURGMAIL@gmail.com"

    recipients = email

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
Attached is your invoice dated {date}.

Thank you for your business.
        
PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'\\Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

new_email = send_email(invoice_num='8515', name = name, email = email, date = date, directory = directory)

filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515
for filename in filenames:
    if ('xlsx' and 'Order') in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1