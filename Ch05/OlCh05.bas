Attribute VB_Name = "OlCh05"
Option Explicit

' Delete all OLE attachments from this item.
Sub DeleteOleAttachments()
Dim ns As NameSpace
Dim draft_folder As MAPIFolder
Dim mail_item As MailItem
Dim i As Integer
Dim att As Attachment

    ' Find the "Progress Report" mail item.
    Set ns = GetNamespace("MAPI")
    Set draft_folder = ns.GetDefaultFolder(olFolderDrafts)
    Set mail_item = draft_folder.Items("Progress Report")

    ' Examine the attachments.
    For i = mail_item.Attachments.Count To 1 Step -1
        Set att = mail_item.Attachments(i)
        If att.Type = olOLE Then att.Delete
    Next i

    ' Save the changes.
    mail_item.Save
End Sub

' Make a simple text-only mail item.
Sub MakeSimpleMailItem()
Dim new_item As MailItem

    ' Create a new mail item.
    Set new_item = Application.CreateItem(olMailItem)
    new_item.To = "Bob@somewhere.com"
    new_item.Subject = "Simple message"
    new_item.Body = _
        "This is a simple mail message containing only text." & vbCrLf & vbCrLf & _
        "Another example attaches this message as an embedded item."
    new_item.Save
End Sub

' Make a mail item with three different kinds
' of attachments.
Sub MakeMailWithAttachments()
Dim new_item As MailItem
Dim ns As NameSpace
Dim draft_folder As MAPIFolder
Dim attachment_item As MailItem

    ' Create a new mail item.
    Set new_item = Application.CreateItem(olMailItem)
    new_item.To = "Cindy@somewhere.com"
    new_item.Subject = "Kinds of attachments"
    new_item.Body = _
        "This message contains three kinds of Outlook attachments: by value, by reference, and embedded item." & vbCrLf & vbCrLf
    new_item.Attachments.Add _
        "C:\OfficeSmackdown\Src\Ch05\Colors.doc", _
        olByValue
    new_item.Attachments.Add _
        "C:\OfficeSmackdown\Src\Ch05\Food.doc", _
        olByReference

    ' Find the "Simple message" mail item.
    Set ns = Application.GetNamespace("MAPI")
    Set draft_folder = ns.GetDefaultFolder(olFolderDrafts)
    Set attachment_item = draft_folder.Items("Simple message")

    ' Attach this item as an embedded item.
    new_item.Attachments.Add _
        attachment_item, _
        olEmbeddeditem
    new_item.Save
End Sub

' Make a mail item with a "by reference"
' attachment in a specific position.
Sub MakePositionedAttachment()
Dim new_item As MailItem
Dim txt1 As String
Dim txt2 As String

    ' Create a new mail item.
    Set new_item = Application.CreateItem(olMailItem)
    new_item.To = "Alice@somewhere.com"
    new_item.Subject = "1Q Sales Figures"
    txt1 = _
        "Alice," & vbCr & vbCr & _
        "Here are the final sales figures for the first quarter." & _
        vbCr & "X"
    txt2 = _
        vbCr & _
        "I hope you find them useful." & vbCr & vbCr & _
        "Rod"
    new_item.Body = txt1 & txt2
    new_item.Save

    ' Add an attachment.
    new_item.Attachments.Add _
        "C:\OfficeSmackdown\Src\Ch05\Sales.xls", _
        olByReference, Len(txt1)
    new_item.Save
End Sub

Sub test()
Dim ns As NameSpace
Dim draft_folder As MAPIFolder
Dim mail_item As MailItem
Dim att As Attachment

    ' Find the "1Q Sales Figures" mail item.
    Set ns = GetNamespace("MAPI")
    Set draft_folder = ns.GetDefaultFolder(olFolderDrafts)
    Set mail_item = draft_folder.Items("OLE Test")
    Set att = mail_item.Attachments(1)
End Sub
