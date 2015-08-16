# Word-Add-in-JavaScript-InvoiceManager
This sample shows how to create a task pane app for Office that manages invoices in Word.


Apps for Office: Create an invoice manager  




Summary: This sample shows how to create a task pane app for Office that manages invoices in Word 2013.

##Description


This sample loads order data into an invoice form in Word 2013. It writes customer data to a set of custom XML parts that are bound to content controls within a Word document. Based on user input, it populates forms in the document with customer and order information. To simplify this sample, the order data is stored in the same JavaScript file that creates the app for Office. However, in a real application, that data could come from a data source anywhere on the web.

The JavaScript code in the Home.js file includes a function for the initialize event, which waits for the DOM to load, gets a reference to the current document, and then calls two other functions. The first of these, setupMyOrders, creates an array to hold the order data.

The second function, initializeOrder, does most of the important work. When the Populate button is chosen, this function first calls the  getByNamespaceAsync method of the  CustomXmlParts object to determine whether the packing slip form is already populated. If it is, the function calls the  deleteAysnc method of the  CustomXmlPart object to delete the existing data in the form. Then it calls the  addAsync method of the  CustomXmlParts object to repopulate the form with the selected data.

##Prerequisites


This sample requires the following:

•Word 2013.


•Visual Studio 2012; App for Office 2013 project template.


•Internet Explorer 9 or Internet Explorer 10.


•Basic familiarity with JavaScript and HTML.


##Key components


The Apps for Office: Create an Invoice Manager sample is created by the InvoiceManager solution, which contains the following projects and important files:

•The InvoiceManager project, including the following files:

•InvoiceManager.xml manifest file


•Packing Slip Document.docx file



•The InvoiceManagerWeb project, including the following files:

•Home.html file


•Home.js file



##Configure the sample


No additional configuration is necessary.

##Build the sample


Choose the F5 key in Visual Studio 2012 to build and deploy the app and open it in Word 2013.

Run and test the sample


1.Open the InvoiceManager.sln file in Visual Studio 2012.


2.Choose the F5 key in Visual Studio 2012 to build and deploy the app.


3.In the app task pane, select an order in the Order ID drop-down list.


4.Choose Populate to populate the forms in the Word document with information from the selected order.


You can view a list of the custom XML parts in a document by opening the XML Mapping pane in Word (Developer tab).

##Troubleshooting


If the app fails to respond as described, try reloading it. (In the task pane, choose the down arrow, and then choose Reload.)

##Change log


Third release.

##Related content


•JavaScript API for Office 


