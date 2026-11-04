required_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:
        directory = select_file("Please select a file in the folder containing your files.")

        directory = directory.split('/')
        new_directory = directory[:-1]
        directory = '\\'.join(new_directory)

        if directory[-1] != '\\':
            directory = directory + '\\'


        fileindir = os.listdir(directory)
        if 'ProductData.xlsx' and 'MasterInvoice.pdf' and 'Order1.xlsx' and 'Order2.xlsx' and 'Order3.xlsx' and 'Order4.xlsx' in fileindir:
            break
        else: 
            show_message("Error", "The folder is missing required files. Please try again.")
        

print(directory)



# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}


descriptions_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
for index, row in descriptions_df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']


prices_df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
for index, row in prices_df.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']


product_descriptions   

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    invoice_num = int(invoice_num)
    df = pd.read_excel(directory + filename)

    template = openpyxl.load_workbook(filename)
    selected_sheet = template['OrderForm']


    order_data = {}

    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    Company = selected_sheet['C15'].value
    date = selected_sheet['C9'].value
    address1 = selected_sheet['C16'].value 
    address2 = str(selected_sheet['C17'].value) + ', ' + str(selected_sheet['C18'].value) + ' ' + str(selected_sheet['C19'].value)

    order_data['InvoiceNum'] = invoice_num

    date = date.strftime("%m/%d/%Y")

    order_data['Date'] = date
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = Company
    order_data['CustomerAddress1'] = address1
    order_data['CustomerAddress2'] = address2
    order_data['Email'] = email

    increment = 24
    dnum = 0
    total = 0

    while True:
        unitno = selected_sheet[f'B{increment}'].value

        if unitno is None:
            break
        else:
            Description = product_descriptions[unitno]
            Price = product_prices[unitno]
            Quantity = selected_sheet[f'C{increment}'].value

            Amount = Price * Quantity
            total = total + Amount

            order_data[f'Description{dnum}'] = Description
            order_data[f'Quantity{dnum}'] = Quantity
            order_data[f'Price{dnum}'] = f'${Price:,.2f}'
            order_data[f'Amount{dnum}'] = f'${Amount:,.2f}'

            increment = increment + 1
            dnum = dnum + 1

        order_data['Total'] = f'${total:,.2f}'
            

    return order_data, name, email, date




order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)

print(order_data)
print(name)
print(email)
print(date)

df1 = get_pdf('MasterInvoice.pdf')
df1


# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):

    original_pdf = directory + 'MasterInvoice.pdf'
    new_pdf = directory + 'Invoice ' + str(invoice_num) + '.pdf'

    create_pdf(original_pdf, new_pdf, order_data)


create_invoice(8515, order_data, directory)


# INSERT ANSWER FOR QUESTION 5 HERE


gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):


    recipients = [email]

    subject = 'Your Order from PythonTech Co. - Invoice ' + str(invoice_num)

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    invoice_file = directory + 'Invoice ' + str(invoice_num) + '.pdf'

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    
    message = add_attachment(invoice_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


send_email(8515, 'Thomas Holt', 'sethfranks1019@gmail.com', '11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)

invoice_num = 8515

for filename in filenames:

    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)

        invoice_num = invoice_num + 1