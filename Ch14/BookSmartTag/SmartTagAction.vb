Imports SmartTagLib
Imports System.Runtime.InteropServices

' Class to provide actions.
<ProgId("BookSmartTag.Action"), _
    GuidAttribute("B28B112D-38FE-427f-847D-FFC4C82B2301"), _
    ComVisible(True)> _
Public Class SmartTagAction
    Implements ISmartTagAction

    ' Methods that identify the action class.

    ' The action class's short name.
    Public ReadOnly Property Name(ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagAction.Name
        Get
            Return "Book Smart Tag Action Class"
        End Get
    End Property

    ' Longer description.
    Public ReadOnly Property Desc(ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagAction.Desc
        Get
            Return "Offer replacements for select book synonyms"
        End Get
    End Property

    ' Return the smart tag's prog ID.
    Public ReadOnly Property ProgId() As String Implements SmartTagLib.ISmartTagAction.ProgId
        Get
            Return "BookSmartTag.SmartTagAction"
        End Get
    End Property

    ' Methods that describe the smart tag types supported.

    ' Caption displayed for the action menu.
    Public ReadOnly Property SmartTagCaption(ByVal SmartTagID As Integer, ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagAction.SmartTagCaption
        Get
            Return "Office Smackdown Actions"
        End Get
    End Property

    ' Number of smart tag types.
    Public ReadOnly Property SmartTagCount() As Integer Implements SmartTagLib.ISmartTagAction.SmartTagCount
        Get
            Return 1
        End Get
    End Property

    ' Smart tag type names in the format URI#TagName.
    Public ReadOnly Property SmartTagName(ByVal SmartTagID As Integer) As String Implements SmartTagLib.ISmartTagAction.SmartTagName
        Get
            Return "http://www.vb-helper.com#BookSmartTag"
        End Get
    End Property

    ' Methods that define the supported verbs.

    ' Number of verbs supported.
    Public ReadOnly Property VerbCount(ByVal SmartTagName As String) As Integer Implements SmartTagLib.ISmartTagAction.VerbCount
        Get
            Return 2
        End Get
    End Property

    ' Unique ID for this verb.
    Public ReadOnly Property VerbID(ByVal SmartTagName As String, ByVal VerbIndex As Integer) As Integer Implements SmartTagLib.ISmartTagAction.VerbID
        Get
            Return VerbIndex
        End Get
    End Property

    ' Caption for this verb.
    Public ReadOnly Property VerbCaptionFromID(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagAction.VerbCaptionFromID
        Get
            Select Case VerbID
                Case 1
                    Return "Change text to Office Smackdown"
                Case 2
                    Return "Open Office Smackdown Web site"
            End Select
        End Get
    End Property

    ' The name for this verb.
    Public ReadOnly Property VerbNameFromID(ByVal VerbID As Integer) As String Implements SmartTagLib.ISmartTagAction.VerbNameFromID
        Get
            Select Case VerbID
                Case 1
                    Return "changeToOfficeSmackdown"
                Case 2
                    Return "openOfficeSmackdownWebSite"
            End Select
        End Get
    End Property

    ' Do whatever's appropriate.
    Public Sub InvokeVerb(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal Target As Object, ByVal Properties As SmartTagLib.ISmartTagProperties, ByVal [Text] As String, ByVal Xml As String) Implements SmartTagLib.ISmartTagAction.InvokeVerb
        Select Case VerbID
            Case 1  ' Change text to "Office Smackdown."
                If ApplicationName.StartsWith("Word.Application") Then
                    Target.Text = "Office Smackdown"
                ElseIf ApplicationName.StartsWith("Excel.Application") Then
                    Target.Value = "Office Smackdown"
                ElseIf ApplicationName.StartsWith("PowerPoint.Application") Then
                    Target.Text = "Office Smackdown"
                End If

            Case 2  ' Open Ofice Smackdown Web site.
                Dim browser As Object
                browser = CreateObject("InternetExplorer.Application")
                browser.Navigate2("http://www.vb-helper.com/office.htm")
                browser.Visible = True
        End Select
    End Sub
End Class ' End SmartTagAction.
