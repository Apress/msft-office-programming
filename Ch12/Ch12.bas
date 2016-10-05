Attribute VB_Name = "Ch12"
Option Explicit

' List all open Explorers and their current contents.
Sub ListExplorers()
Dim the_explorer As Explorer
Dim the_item As Object

    For Each the_explorer In Application.Explorers
        Debug.Print "***** " & the_explorer.Caption
        Debug.Print "Current Folder: " & the_explorer.CurrentFolder.Name
        Debug.Print Format$(the_explorer.CurrentFolder.Items.Count) & " Items:"
        For Each the_item In the_explorer.CurrentFolder.Items
            Debug.Print "    " & TypeName(the_item) & _
                ": " & the_item.Subject
        Next the_item
    Next the_explorer
End Sub
' Search Drafts for mail messages
' with subjects containing "attach".
Sub SearchForAttach()
    Application.AdvancedSearch _
        Scope:="Drafts", _
        Filter:="urn:schemas:mailheader:subject LIKE '%Attach%'", _
        Tag:="attach"
End Sub
' Search Calendar for items
' with subjects containing "stuff".
Sub SearchForStuff()
    Application.AdvancedSearch _
        Scope:="Calendar", _
        Filter:="urn:schemas:mailheader:subject LIKE '%stuff%'", _
        Tag:="stuff"
End Sub
' Copy some test files into the Notes and Inbox folders.
Sub CopyFiles()
Dim file_names As Variant
Dim i As Integer

    ' Define the file names.
    file_names = Array( _
        "C:\OfficeSmackdown\Src\Ch12\CodeFragment.txt", _
        "C:\OfficeSmackdown\Src\Ch12\TextFragment.doc", _
        "C:\OfficeSmackdown\Src\Ch12\Test.html", _
        "C:\OfficeSmackdown\Src\Ch12\dog.jpg")

    ' Copy the files.
    For i = LBound(file_names) To UBound(file_names)
        Application.CopyFile file_names(i), "Notes"
        Application.CopyFile file_names(i), "Inbox"
    Next i
End Sub
' Make a new 90 minute appointment item
' for 1:00 PM tomorrow.
Sub MakeAppointmentTomorrow()
Dim appt As AppointmentItem

    Set appt = Application.CreateItem(olAppointmentItem)
    With appt
        .Subject = "Conference call re. Office Smackdown"
        .Body = "Editor will call."
        .Start = DateAdd("d", 1, Date) + #1:00:00 PM#
        .Duration = 90
        .Importance = olImportanceHigh
        .Location = "Office"
        .ReminderMinutesBeforeStart = 30
        .Save
    End With
End Sub
' Make a new Post item in Inbox.
Sub MakeAppointmentTomorrowInInbox()
Dim inbox_folder As MAPIFolder
Dim post_item As PostItem

    ' Get the Inbox folder from the session's MAPI namespace.
    Set inbox_folder = Session.GetDefaultFolder(olFolderInbox)

    ' Create the new item.
    Set post_item = inbox_folder.Items.Add(olPostItem)
    With post_item
        .Subject = "Need stamps"
        .Body = "This is a simple Post item."
        .ExpiryTime = DateAdd("d", 2, Now)
        .Importance = olImportanceHigh
        .Save
    End With
End Sub

' Make a new Post item in Inbox.
Sub xMakeAppointmentTomorrowInInbox()
Dim inbox_folder As MAPIFolder
Dim post_item As ContactItem

    ' Get the Inbox folder from the session's MAPI namespace.
    Set inbox_folder = Session.GetDefaultFolder(olFolderCalendar)

    ' Create the new item.
    Set post_item = inbox_folder.Items.Add(olContactItem)
    With post_item
        .LastName = "Stephens"
        .FirstName = "Cobe"
        .Body = "A cat"
        .Save
    End With
End Sub
' Close Outlook discarding any unsaved changes.
Sub CloseOutlook()
    ' Close all open Inspectors.
    Do While Inspectors.Count > 0
        Inspectors(1).Close olDiscard
    Loop

    ' Close Outlook.
    Application.Quit
End Sub
' Display the subjects of unread messages.
Sub ShowUnreadMessages()
Dim inbox_folder As MAPIFolder
Dim itm As Object

    Set inbox_folder = Session.GetDefaultFolder(olFolderInbox)
    For Each itm In inbox_folder.Items
        If itm.UnRead Then
            MsgBox itm.Subject
        End If
    Next itm
End Sub
' Add a copyright statement at the end of the current item.
Sub AddCopyright()
    ActiveInspector.CurrentItem.Body = _
        ActiveInspector.CurrentItem.Body & vbCrLf & _
        "Copyright " & Year(Date) & " S. Nob Esq."
End Sub
' Display the top-level folders and their descendants.
Sub DisplayAllFolders()
Dim child_folder As MAPIFolder

    For Each child_folder In Session.Folders
        DisplayFolder child_folder, 0
    Next child_folder
End Sub
' Display the folder hierarchy within this folder.
Sub DisplayFolder(ByVal start_folder As MAPIFolder, ByVal level As Integer)
Dim child_folder As MAPIFolder

    Debug.Print Space$(level) & start_folder.Name
    For Each child_folder In start_folder.Folders
        DisplayFolder child_folder, level + 4
    Next child_folder
End Sub
' Make a sub-folder inside Inbox named Spam.
Sub MakeInboxSpamFolder()
Dim inbox_folder As MAPIFolder

    Set inbox_folder = Session.GetDefaultFolder(olFolderInbox)
    inbox_folder.Folders.Add "Spam", olFolderInbox
End Sub
' Delete the Inbox/Spam folder.
Sub DeleteInboxSpamFolder()
Dim inbox_folder As MAPIFolder
Dim spam_folder As MAPIFolder

    Set inbox_folder = Session.GetDefaultFolder(olFolderInbox)
    Set spam_folder = inbox_folder.Folders("Spam")
    spam_folder.Delete
End Sub
' Display the Deleted Items folder's contents.
Sub DisplayDeletedFolderContents()
Dim deleted_folder As MAPIFolder

    ' Find the folder.
    Set deleted_folder = Session.GetDefaultFolder(olFolderDeletedItems)

    ' Display the folder's contents.
    DisplayFolderContents deleted_folder, 0
End Sub
' Display the folder, its items, and its subfolders.
Sub DisplayFolderContents(ByVal start_folder As MAPIFolder, ByVal level As Integer)
Dim obj As Object
Dim child_folder As MAPIFolder

    ' Display the folder's name.
    Debug.Print Space$(level) & "Folder: " & start_folder.Name

    ' Display the items in this folder.
    For Each obj In start_folder.Items
        Debug.Print Space$(level + 4) & "[" & obj.Subject & "]"
    Next obj

    ' Display the subfolders.
    For Each child_folder In start_folder.Folders
        DisplayFolderContents child_folder, level + 4
    Next child_folder
End Sub
' Add a recipient to every mail message in the
' ProgressReports folder.
Sub AddRecipient()
Dim pr_folder As MAPIFolder
Dim obj As Object
Dim mail_item As MailItem

    Set pr_folder = Session.Folders("Personal Folders").Folders("ProgressReports")
    For Each obj In pr_folder.Items
        ' See if this is a MailItem.
        If obj.Class = olMail Then
            Set mail_item = obj
            mail_item.Recipients.Add "121ProjectHistory@nowhere.com"
            mail_item.Save
        End If
    Next obj
End Sub
' List the items selected in the active Explorer.
Sub ListSelectedItems()
Dim obj As Object

    For Each obj In Application.ActiveExplorer.Selection
        Debug.Print obj.Subject
    Next obj
End Sub
' Make a new view in the Inbox.
Sub CreateView()
Dim inbox_views As Views
Dim new_view As View

    Set inbox_views = Application.Session.GetDefaultFolder(olFolderInbox).Views
    Set new_view = inbox_views.Add( _
        Name:="Sort By Subject", _
        ViewType:=olTableView, _
        SaveOption:=olViewSaveOptionThisFolderEveryone)
End Sub
' Display the XML code defining the Sort By Subject view.
Sub ListViewXML()
Dim inbox_views As Views

    Set inbox_views = Application.Session.GetDefaultFolder(olFolderInbox).Views
    Debug.Print inbox_views("Sort By Subject").xml
End Sub
' List available address lists and their entries.
Sub ListAddresses()
Dim address_list As AddressList
Dim address_entry As AddressEntry

    For Each address_list In Session.AddressLists
        Debug.Print "***** " & address_list.Name & " *****"
        For Each address_entry In address_list.AddressEntries
            Debug.Print address_entry.Name & _
                " (" & address_entry.Address & ")"
        Next address_entry
    Next address_list
End Sub
' List an item's actions.
Sub ListActions()
Dim draft_item As MailItem
Dim an_action As Action

    Set draft_item = Session.GetDefaultFolder(olFolderDrafts).Items(1)
    For Each an_action In draft_item.Actions
        Debug.Print an_action.Name
    Next an_action
End Sub
' Invoke an item's Reply to All action.
Sub ReplyToAllAction()
Dim draft_item As MailItem
Dim reply_action As Action
Dim reply_item As MailItem

    Set draft_item = Session.GetDefaultFolder(olFolderDrafts).Items("1Q Sales Figures")
    Set reply_action = draft_item.Actions("Reply to All")
    reply_action.ReplyStyle = olIndentOriginalText
    Set reply_item = reply_action.Execute()
    reply_item.Subject = "Re: " & reply_item.Subject
    reply_item.Save
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

    ' Find the "Simple messaage" mail item.
    Set ns = Application.GetNamespace("MAPI")
    Set draft_folder = ns.GetDefaultFolder(olFolderDrafts)
    Set attachment_item = draft_folder.Items("Simple message")

    ' Attach this item as an embedded item.
    new_item.Attachments.Add _
        attachment_item, _
        olEmbeddeditem
    new_item.Save
End Sub
' Make a project review appointment that happens
' every Tuesday at 4:00 pm.
Sub MakeProjectReviewAppointment()
Dim appt As AppointmentItem
Dim recurrence_pattern As RecurrencePattern

    ' Make the new appointment.
    Set appt = Application.CreateItem(olAppointmentItem)
    appt.Subject = "Weekly project review"
    appt.Body = "Review the week's progress, outstanding issues, etc."

    ' Make the recurrence.
    Set recurrence_pattern = appt.GetRecurrencePattern()
    recurrence_pattern.DayOfWeekMask = olTuesday
    recurrence_pattern.StartTime = "4:00PM"
    recurrence_pattern.Duration = 60
    recurrence_pattern.NoEndDate = True

    ' Save the appointment.
    appt.Save
End Sub
' List reminders.
Sub ViewReminderInfo()
Dim a_reminder As Reminder

    For Each a_reminder In Application.Reminders
        Debug.Print "*****"
        Debug.Print a_reminder.NextReminderDate & _
            " (" & a_reminder.Item.Start & ")"
        Debug.Print a_reminder.Caption
        Debug.Print a_reminder.Item.Body
    Next a_reminder
End Sub
' Add the To field to the Inbox's "Subject and From" view.
Sub FromToField()
Dim inbox_views As Views
Dim xml As String
Dim to_field As String
Dim pos As Integer

    ' Get the view's current XML.
    Set inbox_views = Application.Session.GetDefaultFolder(olFolderInbox).Views
    xml = inbox_views("Subject and From").xml

    ' See if the To field is already present.
    If InStr(xml, "<prop>urn:schemas:httpmail:displayto</prop>") > 0 _
        Then Exit Sub

    ' Find the end of the last column.
    pos = InStrRev(xml, "</column>") + Len("</column>")

    ' Add the From field.
    to_field = vbCrLf & _
"    <column>" & vbCrLf & _
"        <heading>To</heading>" & vbCrLf & _
"        <prop>urn:schemas:httpmail:displayto</prop>" & vbCrLf & _
"        <type>string</type>" & vbCrLf & _
"        <width>47</width>" & vbCrLf & _
"        <style>text-align:left;padding-left:3px</style>" & vbCrLf & _
"    </column>"

    xml = Left$(xml, pos) & to_field & Mid$(xml, pos)

    ' Set the view's new XML value.
    inbox_views("Subject and From").xml = xml
End Sub
' Make a mail item with To, Cc, and Bcc recipients.
Sub MakeMailRecipients()
Dim mail_item As MailItem
Dim recip As Recipient

    ' Create a new mail item.
    Set mail_item = Application.CreateItem(olMailItem)

    ' Add some recipients.
    ' This one's Type defaults to olTo.
    Set recip = mail_item.Recipients.Add("Cindy@somewhere.com")

    Set recip = mail_item.Recipients.Add("Bob@somewhere.com")
    recip.Type = olCC

    Set recip = mail_item.Recipients.Add("Sandy@somewhere.com")
    recip.Type = olCC

    Set recip = mail_item.Recipients.Add("MikeH@somewhere.com")
    recip.Type = olBCC

    mail_item.Subject = "Recipients"
    mail_item.Body = _
        "This message has To, Cc, and Bcc recipients."

    mail_item.Save
End Sub
' Create a meeting appointment.
Sub CreateMeeting()
Dim appt_item As AppointmentItem
Dim attendee As Recipient

    ' Create the appointment.
    Set appt_item = Application.CreateItem(olAppointmentItem)

    ' Set some descriptive properties.
    appt_item.Subject = "Project Review"
    appt_item.Location = "Rod's office"
    appt_item.Start = #8/20/2003 4:00:00 PM#
    appt_item.Duration = 45
    appt_item.ReminderMinutesBeforeStart = 15

    ' Make it a meeting.
    appt_item.MeetingStatus = olMeeting

    ' Specify attendees.
    appt_item.Recipients.Add("Rod Stephens").Type = olOrganizer
    appt_item.Recipients.Add("Alice Archer").Type = olRequired
    appt_item.Recipients.Add("Bill Blah").Type = olRequired
    appt_item.Recipients.Add("Cindy Carter").Type = olOptional
    appt_item.Recipients.Add("David Dart").Type = olResource

    ' Send the appointment to the recipients.
    appt_item.Send
End Sub
