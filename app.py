# Insert your python code here
while True:
    directory = single_input("Please enter the file path to your folder.")
    directory = directory.replace('"','')
    if directory.endswith('/') != True:
        directory = directory +'/'
    try:
        filenames = os.listdir(directory)
        if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
            break
        else:
            show_message("Error", "The file path you entered is incorrect. Please try again.")
    except:
        show_message("Error", "The file path you entered is incorrect. Please try again.")
print(directory)





product_descriptions = {}
product_prices = {}

ProductData = 'ProductData.xlsx'
full_path = os.path.join(directory, ProductData)

dfd = pd.read_excel(full_path, sheet_name='Descriptions')
for index, row in dfd.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

print(product_descriptions)

dfp = pd.read_excel(full_path, sheet_name = 'Prices')
for index, row in dfp.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

print(product_prices)






import datetime

def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']

    order_data_dict = {}

    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    email = selected_sheet['C13'].value
    name = selected_sheet['C14'].value
    company = selected_sheet['C15'].value
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value

    order_data_dict['InvoiceNum'] = int(invoice_num)
    order_data_dict['Date'] = date
    order_data_dict['CustomerName'] = name
    order_data_dict['CustomerCompany'] = company
    order_data_dict['CustomerAddress1'] = address
    order_data_dict['CustomerAddress2'] = f'{city}, {state} {zipcode}'

    total = 0
    row = 24
    item = 0

    while True:
        unit_num = selected_sheet[f'B{row}'].value
        qty = selected_sheet[f'C{row}'].value
        if unit_num is None:
            break
        desc = str(product_descriptions[unit_num])
        price = product_prices[unit_num]

        amount = price * qty
        total = total + amount

        order_data_dict[f'Description{item}'] = desc
        order_data_dict[f'Price{item}'] = f'${price:,.2f}'
        order_data_dict[f'Quantity{item}'] = str(qty)
        order_data_dict[f'Amount{item}'] = f'${amount:,.2f}'
   
        row = row + 1
        item = item + 1

    order_data_dict['Total'] = f'${total:,.2f}'
    return order_data_dict, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(f'{order_data, name, email, date}')






def create_invoice(invoice_num,order_data,directory):
    input = os.path.join(directory,'MasterInvoice.pdf')
    output = os.path.join(directory, f'Invoice {invoice_num}.pdf')
    create_pdf(input, output, order_data)
create_invoice(8515, order_data, directory)







gpass = open('gmail_password.txt','w')
gpass.write('ypsydfrofjbmudql')
gpass.close()

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email (invoice_num, name, email, date, directory):
    gmail_user = "emmehiggins1107@gmail.com"

    recipients = email

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372"""

    # Initialize the email message
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment
    filename = f'Invoice {invoice_num}.pdf'
    path_to_file = os.path.join(directory,filename)
    message = add_attachment(path_to_file, message)  

    # Log in to the Gmail server
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email
    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,'Thomas Holt','emmehiggins1107@gmail.com','11/13/2023',directory)






filenames = os.listdir(directory)
filenames.sort()
print(filenames)

number = 8515

for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(number, file, directory)
        create_invoice(number, order_data, directory)
        send_email(number, name, 'emmehiggins1107@gmail.com', date, directory)
        number = number + 1