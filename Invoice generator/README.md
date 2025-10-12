**Invoice generator**
Purpose of this tool is to create invoices, store them and send to customers

The tool is build in excel (version: 2509 Build 16.0.19231.20138) with Virtual Basic for Application (VBA). Created macros are assigned to buttons for better user experience. Each macro inform the user about results

**How to use it**
1. Download the excel file
2. Open the file and enable macros
3. Adjust company details
4. Populate "Customers" tab with data of your customers
5. Create an invoice by:
    5.1 Fill PO number
    5.2 Adjust invoice date 
    5.3 Select customer from the dropdown list in cell B9
    5.4 Fill information about produtcs/services, their quantity, price and VAT rate
    5.5 (optional) if inserted data is invalid, please use "create a new invoice" to start the process over
    5.6 Click "Save and export to PDF" button
6. Send invoices to customers by clicking "Send invoices to customers" button

**Buttons functionality**
Save and export to PDF - Creates a record of the invoice in "Record of Invoices" tab and exports it to PDF file which is stored in ./invoices/ PDF files are name as follow: "(invoice number) - (customer name)" for example " 1001 - Rommold GmbH.pdf". After that it clears the template, so a new invoice can be created

Send invoices to customers - Sends all unsent invoices to customers with use of Outlook. Destination address is picked from customer email address in "Customers" tab

Create a new invoice - cleares the invoice template