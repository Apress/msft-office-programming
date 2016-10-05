Imports SmartTagLib
Imports System.Runtime.InteropServices

<ProgId("FlavorsSmartTag.SmartTagAction"), _
    GuidAttribute("87BDE0FF-248E-4b9e-8EE5-C2EF2579821E"), _
    ComVisible(True)> _
Public Class SmartTagAction
    Implements ISmartTagAction
    Implements ISmartTagAction2

    ' ************************
    ' ISmartTagAction methods.
    ' ************************
    ' The class's short name.
    Public ReadOnly Property Name(ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagAction.Name
        Get
            Return "Flavor Smart Tag Action Class"
        End Get
    End Property

    ' Longer description.
    Public ReadOnly Property Desc(ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagAction.Desc
        Get
            Return "Provide flavor verbs"
        End Get
    End Property

    ' ProgId.
    Public ReadOnly Property ProgId() As String Implements SmartTagLib.ISmartTagAction.ProgId
        Get
            Return "FlavorsSmartTag.SmartTagAction"
        End Get
    End Property

    ' Number of smart tags supported by this class.
    Public ReadOnly Property SmartTagCount() As Integer Implements SmartTagLib.ISmartTagAction.SmartTagCount
        Get
            Return 1
        End Get
    End Property

    ' Menu caption.
    Public ReadOnly Property SmartTagCaption(ByVal SmartTagID As Integer, ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagAction.SmartTagCaption
        Get
            Return "Flavors"
        End Get
    End Property

    ' Smart tag name.
    Public ReadOnly Property SmartTagName(ByVal SmartTagID As Integer) As String Implements SmartTagLib.ISmartTagAction.SmartTagName
        Get
            Return "http://www.vb-helper.com#FlavorsSmartTag"
        End Get
    End Property

    ' Number of verbs supported.
    Public ReadOnly Property VerbCount(ByVal SmartTagName As String) As Integer Implements SmartTagLib.ISmartTagAction.VerbCount
        Get
            Return 4
        End Get
    End Property

    ' Verb IDs for verb numbers.
    Public ReadOnly Property VerbID(ByVal SmartTagName As String, ByVal VerbIndex As Integer) As Integer Implements SmartTagLib.ISmartTagAction.VerbID
        Get
            Return VerbIndex
        End Get
    End Property

    ' Verb names.
    Public ReadOnly Property VerbNameFromID(ByVal VerbID As Integer) As String Implements SmartTagLib.ISmartTagAction.VerbNameFromID
        Get
            Select Case VerbID
                Case 1
                    Return "changeToChocolate"
                Case 2
                    Return "changeToVanilla"
                Case 3
                    Return "changeToStrawberry"
                Case 4
                    Return "goToWebSite"
            End Select
        End Get
    End Property

    ' Verb menu captions.
    ' This is provided for backwards compatibility and 
    ' is used only if the ISmartTagAction2 is not supported.
    Public ReadOnly Property VerbCaptionFromID(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagAction.VerbCaptionFromID
        Get
            Select Case VerbID
                Case 1
                    Return "Change To CHOCOLATE"
                Case 2
                    Return "Change To VANILLA"
                Case 3
                    Return "Change To STRAWBERRY"
                Case 4
                    Return "Go To WEB SITE"
            End Select
        End Get
    End Property

    ' Invoke the verb.
    ' This is provided for backwards compatibility and 
    ' is used only if the ISmartTagAction2 is not supported.
    Public Sub InvokeVerb(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal Target As Object, ByVal Properties As SmartTagLib.ISmartTagProperties, ByVal recognized_text As String, ByVal Xml As String) Implements SmartTagLib.ISmartTagAction.InvokeVerb
        If VerbID <= 3 Then
            ' Make a replacement.
            PerformReplacement(VerbID, ApplicationName, Target, recognized_text)
        Else
            ' Go to the Web site.
            Dim browser As Object
            browser = CreateObject("InternetExplorer.Application")
            browser.Navigate2("http://www.vb-helper.com/office.htm")
            browser.Visible = True
        End If
    End Sub


    ' *************************
    ' ISmartTagAction2 methods.
    ' *************************
    ' Here you can do things like save the host
    ' application's name.
    Public Sub SmartTagInitialize(ByVal ApplicationName As String) Implements SmartTagLib.ISmartTagAction2.SmartTagInitialize

    End Sub

    ' Captions for the verbs.
    Public ReadOnly Property VerbCaptionFromID2(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer, ByVal Properties As SmartTagLib.ISmartTagProperties, ByVal recognized_text As String, ByVal Xml As String, ByVal Target As Object) As String Implements SmartTagLib.ISmartTagAction2.VerbCaptionFromID2
        Get
            Select Case VerbID
                Case 1
                    Return "Replace with...///Chocolate"
                Case 2
                    Return "Replace with...///Vanilla"
                Case 3
                    Return "Replace with...///Strawberry"
                Case 4
                    Return "Go To Web Site"
            End Select
        End Get
    End Property

    ' Return True if the caption may change each time 
    ' the menu is displayed.
    Public ReadOnly Property IsCaptionDynamic(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer) As Boolean Implements SmartTagLib.ISmartTagAction2.IsCaptionDynamic
        Get
            Return False
        End Get
    End Property

    ' Return True if you want the host to display
    ' the smart tag indicator.
    Public ReadOnly Property ShowSmartTagIndicator(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal LocaleID As Integer) As Boolean Implements SmartTagLib.ISmartTagAction2.ShowSmartTagIndicator
        Get
            Return True
        End Get
    End Property

    ' Perform the appropriate action.
    Public Sub InvokeVerb2(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal Target As Object, ByVal Properties As SmartTagLib.ISmartTagProperties, ByVal recognized_text As String, ByVal Xml As String, ByVal LocaleID As Integer) Implements SmartTagLib.ISmartTagAction2.InvokeVerb2
        If VerbID <= 3 Then
            ' Make a replacement.
            PerformReplacement(VerbID, ApplicationName, Target, recognized_text)
        Else
            ' Go to the Web site.
            Dim browser As Object
            browser = CreateObject("InternetExplorer.Application")
            browser.Navigate2("http://www.vb-helper.com/office.htm")
            browser.Visible = True
        End If
    End Sub

    ' Replace the recognized text.
    Public Sub PerformReplacement(ByVal VerbID As Integer, ByVal ApplicationName As String, ByVal Target As Object, ByVal recognized_text As String)
        Dim new_text As String

        ' Figure out what to replace the text with.
        Select Case VerbID
            Case 1
                new_text = "chocolate"
            Case 2
                new_text = "vanilla"
            Case 3
                new_text = "strawberry"
        End Select

        ' Set the proper case.
        If recognized_text = recognized_text.ToLower() Then
            ' Lower case.
            new_text = new_text.ToLower()
        ElseIf recognized_text = recognized_text.ToUpper() Then
            ' Upper case.
            new_text = new_text.ToUpper()
        Else
            ' Mixed case.
            new_text = StrConv(new_text, VbStrConv.ProperCase)
        End If

        ' Replace the text for different hosts.
        If ApplicationName.StartsWith("Word.Application") Then
            Target.Text = new_text
        ElseIf ApplicationName.StartsWith("Excel.Application") Then
            Target.Value = new_text
        ElseIf ApplicationName.StartsWith("PowerPoint.Application") Then
            Target.Text = new_text
        End If
    End Sub
End Class
