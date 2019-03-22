# OfficeCustomPropertyBug
Sample application to demonstrate the office custom property save bug
The solution contains two projects. First project contains the manifest for the addin. 
The second project contains an aspx web application. Deploy the web application to your localhost
Deploy the addin manifest to your Office 365 instance.
Open an office 365 word document in your Word Desktop client. 
Open the Addin named "CodeCrasher Office". It will load an aspx with a button "Add custom property and Save office document"
When you hit the button the code adds a custom property named "CodeCrasherID" to the word document and the document is saved. 
You can confirm the addition of the custom property by hitting the Word menu File -> Info -> Properties -> Advanced Properties -> Custom Tab
Now close the word document.
Reopen the word document and navigate to the menu File -> Info -> Properties -> Advanced Properties -> Custom Tab
The custom property is missing
You can also try to manually save the document before closing it. Even then the custom property doesnt get saved into the document.
However if you type some content into the document before closing it then the custom property gets saved into the document.
