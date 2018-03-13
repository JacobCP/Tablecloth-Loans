# Tablecloth-Loans
MS Access Desktop App for managing Tablecloth/Centerpiece Loans

# Overview
Database App for a party decor (tablecloths and centerpieces) free-loan organization. 
Implements a comprehensive system to manage:
  * inventory
  * customers 
  * scheduling 
  * reservations
  * pickups
  * returns 
  * availability 
  * fees
  * payments

To facilitate a quicker task flow, App includes the ability to print out picture catalogs of inventory items 
and integrates with off-the-shelf bar code scanners.

App integrates with Outlook and uses automatically populated templates to send, or print out, 
reservation confirmations, pickup and return reminders, and overdue return warnings to customers.

The sample files include real inventory data, 
and randomly generate order/loan information when opening for the first time.
The code that randomly generates the data is in mdlDemo, and is called from the 

System Requirements:

The app was created with Access 2013, but a later version of Access should work as well.

Download links:

Runtime 2016: https://www.microsoft.com/en-ca/download/details.aspx?id=50040

Runtime 2013: https://www.microsoft.com/en-us/download/details.aspx?id=39358 

The catalogs require the font IDAutomationHC39M, which used to be provided for free,
but currently costs $159 for a
https://www.idautomation.com/free-barcode-products/code39-font/single user license.

This font may be an alternative, but it has not been tested.
https://www.barcodesinc.com/free-barcode-font/

Please note:
Place both the front end and back end files in the same folder.
You may need to close and reopen the front end file in order for the links to the back end file to become functional.

Administrative tabs are hidden by default, and can be shown by using the CTRL+G shortcut.

In order to make use of the email features, you must have Microsoft Outlook setup with an active email account.
Because of Microsoft issues, emails may not automatically send from Outlook, until you open outlook and Send/Receive manually.

# Coding/Design Style
I wrote the code over the span of a few months, during which I was learning Microsoft Access.
A lot of the code and design style is therefore inconsistent, and the earlier work often does not follow recommended practices.
The mdlDemo code that populates the database with random order data was written recently, and is a good reference for my current
level of coding skill in VBA and SQL, and functional programming design.
