# Word add-in: Load data into custom XML parts bound to content controls in a Word document

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="summary"></a>
##Summary


This sample loads order data into an invoice form in Word 2013. It writes customer data to a set of custom XML parts that are bound to content controls within a Word document. Based on user input, it populates forms in the document with customer and order information. To simplify this sample, the order data is stored in the same JavaScript file that creates the app for Office. However, in a real application, that data could come from a data source anywhere on the web.

The JavaScript code in the Home.js file includes a function for the initialize event, which waits for the DOM to load, gets a reference to the current document, and then calls two other functions. The first of these, setupMyOrders, creates an array to hold the order data.

The second function, initializeOrder, does most of the important work. When the Populate button is chosen, this function first calls the  getByNamespaceAsync method of the  CustomXmlParts object to determine whether the packing slip form is already populated. If it is, the function calls the  deleteAysnc method of the  CustomXmlPart object to delete the existing data in the form. Then it calls the  addAsync method of the  CustomXmlParts object to repopulate the form with the selected data.

<a name="prerequisites"></a>
##Prerequisites
This sample requires the following:  

  - Visual Studio 2013 with Update 5 or Visual Studio 2015.  
  - Word 2013 or later
  - Internet Explorer 9 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.
  - One of the following as the default browser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.
  - Familiarity with JavaScript programming and web services.

<a name="components"></a>
##Key components


The Apps for Office: Create an Invoice Manager sample is created by the InvoiceManager solution, which contains the following projects and important files:

•The InvoiceManager project, including the following files:

•InvoiceManager.xml manifest file


•Packing Slip Document.docx file



•The InvoiceManagerWeb project, including the following files:

•Home.html file


•Home.js file

<a name="codedescription"></a>
##Description of the code

<a name="build"></a>
##Build the sample


Choose the F5 key in Visual Studio 2012 to build and deploy the app and open it in Word 2013.

Run and test the sample


1.Open the InvoiceManager.sln file in Visual Studio 2012.


2.Choose the F5 key in Visual Studio 2012 to build and deploy the app.


3.In the app task pane, select an order in the Order ID drop-down list.


4.Choose Populate to populate the forms in the Word document with information from the selected order.


You can view a list of the custom XML parts in a document by opening the XML Mapping pane in Word (Developer tab).

<a name="troubleshooting"></a>
## Troubleshooting

- If the add-in starts with a blank document, ensure that the **Start Document** property of the InvoiceManager project is set to *PackingSlip.docx* and not just to Word.
![](/assets/start_props.png)
- If the add-in does not appear in the task pane, Choose **Insert > My Add-ins >  InvoiceManagerSample**.

<a name="questions"></a>
## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Word-Add-in-JavaScript-InvoiceManager/issues).
- Questions about Office Add-ins development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [Office Add-ins](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Bindings object (JavaScript API for Office)](http://msdn.microsoft.com/library/office/apps/fp160966.aspx)
- [Binding to regions in a document or spreadsheet](http://msdn.microsoft.com/library/office/apps/fp123511(v=office.15).aspx)


## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.


