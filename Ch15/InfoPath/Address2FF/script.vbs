' This file contains functions for data validation and form-level events.
' Because the functions are referenced in the form definition (.xsf) file, 
' it is recommended that you do not modify the name of the function,
' or the name and number of arguments.

' The following line is created by Microsoft Office InfoPath to define the prefixes
' for all the known namespaces in the main XML data file.
' Any modification to the form files made outside of InfoPath
' will not be automatically updated.
'<namespacesDefinition>
XDocument.DOM.setProperty "SelectionNamespaces", "xmlns:my=""http://schemas.microsoft.com/office/infopath/2003/myXSD/2003-08-08T15:13:41"""
'</namespacesDefinition>


'=======
' The following function handler is created by Microsoft Office InfoPath.
' Do not modify the name of the function, or the name and number of arguments.
' This function is associated with the following field or group (XPath): /my:myFields/my:txtZip
' Note: Information in this comment is not updated after the function handler is created.
'=======
Sub msoxd_my_txtZip_OnBeforeChange(eventObj)
' Write your code here
' Warning: ensure that the constraint you are enforcing is compatible with the default value you set for this XML node.

	' Note: eventObj is an object of type DataDOMEvent.
	' Use a RegExp object to verify that the new value
	' looks like a ZIP code.
	Dim reg_exp, matches
	Set reg_exp = New RegExp
	reg_exp.Pattern = "^[0-9]{5}$"
	Set matches = reg_exp.Execute(eventObj.NewValue)
	If matches.Count < 1 Then
	    eventObj.ReturnMessage = "Invalid Zip code format"
        eventObj.ReturnStatus = False
	End If
End Sub
