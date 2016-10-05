Imports SmartTagLib
Imports System.Runtime.InteropServices

<ProgId("FlavorsSmartTag.SmartTagRecognizer"), _
    GuidAttribute("C3274E21-5136-4a36-BBCA-78D8A8D49AC3"), _
    ComVisible(True)> _
Public Class SmartTagRecognizer
    Implements ISmartTagRecognizer
    Implements ISmartTagRecognizer2

    ' ****************************
    ' ISmartTagRecognizer methods.
    ' ****************************
    ' The recognizer's name.
    Public ReadOnly Property Name(ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagRecognizer.Name
        Get
            Return "Flavor Smart Tag Recognizer"
        End Get
    End Property

    ' Longer description.
    Public ReadOnly Property Desc(ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagRecognizer.Desc
        Get
            Return "Recognize the word 'flavor'"
        End Get
    End Property

    ' ProgId.
    Public ReadOnly Property ProgId() As String Implements SmartTagLib.ISmartTagRecognizer.ProgId
        Get
            Return "FlavorsSmartTag.SmartTagRecognizer"
        End Get
    End Property

    ' Number of smart tag types recognized.
    Public ReadOnly Property SmartTagCount() As Integer Implements SmartTagLib.ISmartTagRecognizer.SmartTagCount
        Get
            Return 1
        End Get
    End Property

    ' URL to download actions.
    Public ReadOnly Property SmartTagDownloadURL(ByVal SmartTagID As Integer) As String Implements SmartTagLib.ISmartTagRecognizer.SmartTagDownloadURL
        Get
            Return ""
        End Get
    End Property

    ' The smart tag's name.
    Public ReadOnly Property SmartTagName(ByVal SmartTagID As Integer) As String Implements SmartTagLib.ISmartTagRecognizer.SmartTagName
        Get
            Return "http://www.vb-helper.com#FlavorsSmartTag"
        End Get
    End Property

    ' See if we can recognize anything.
    ' This is provided for backwards compatibility and 
    ' is used only if the ISmartTagRecognizer2 is not supported.
    Public Sub Recognize(ByVal txt As String, ByVal DataType As SmartTagLib.IF_TYPE, ByVal LocaleID As Integer, ByVal RecognizerSite As SmartTagLib.ISmartTagRecognizerSite) Implements SmartTagLib.ISmartTagRecognizer.Recognize
        txt = txt.ToLower()
        If txt.IndexOf("flavor") >= 0 Then
            RecognizerSite.CommitSmartTag( _
                "http://www.vb-helper.com#FlavorsSmartTag", _
                txt.IndexOf("flavor") + 1, _
                6, RecognizerSite.GetNewPropertyBag())
        End If
    End Sub


    ' *****************************
    ' ISmartTagRecognizer2 methods.
    ' *****************************
    ' Here you can do things like save the host
    ' application's name.
    Public Sub SmartTagInitialize(ByVal ApplicationName As String) Implements SmartTagLib.ISmartTagRecognizer2.SmartTagInitialize

    End Sub

    ' Display property pages if provided.
    Public Sub DisplayPropertyPage(ByVal SmartTagID As Integer, ByVal LocaleID As Integer) Implements SmartTagLib.ISmartTagRecognizer2.DisplayPropertyPage

    End Sub

    ' Return True if you can provide property pages.
    Public ReadOnly Property PropertyPage(ByVal SmartTagID As Integer, ByVal LocaleID As Integer) As Boolean Implements SmartTagLib.ISmartTagRecognizer2.PropertyPage
        Get
            Return False
        End Get
    End Property

    ' See if we recognize anything.
    Public Sub Recognize2(ByVal examine_text As String, ByVal DataType As SmartTagLib.IF_TYPE, ByVal LocaleID As Integer, ByVal RecognizerSite2 As SmartTagLib.ISmartTagRecognizerSite2, ByVal ApplicationName As String, ByVal TokenList As SmartTagLib.ISmartTagTokenList) Implements SmartTagLib.ISmartTagRecognizer2.Recognize2
        Dim i As Integer
        For i = 1 To TokenList.Count
            If TokenList.Item(i).Text.ToLower = "flavor" Then
                RecognizerSite2.CommitSmartTag2( _
                    "http://www.vb-helper.com#FlavorsSmartTag", _
                    TokenList.Item(i).Start, _
                    TokenList.Item(i).Length, _
                    RecognizerSite2.GetNewPropertyBag())
            End If
        Next i
    End Sub
End Class
