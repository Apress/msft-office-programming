Imports SmartTagLib
Imports System.Runtime.InteropServices

' Class to recognize terms.
<ProgId("BookSmartTag.SmartTagRecognizer"), _
    GuidAttribute("1B580979-D406-4f1c-B337-D7C8A0E100F0"), _
    ComVisible(True)> _
Public Class SmartTagRecognizer
    Implements ISmartTagRecognizer

    ' Methods that describe the recognizer.

    ' The recognizer's short name.
    Public ReadOnly Property Name(ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagRecognizer.Name
        Get
            Return "Book Smart Tag Recognizer Class"
        End Get
    End Property

    ' Longer description.
    Public ReadOnly Property Desc(ByVal LocaleID As Integer) As String Implements SmartTagLib.ISmartTagRecognizer.Desc
        Get
            Return "Watch for select book synonyms"
        End Get
    End Property

    ' Return the smart tag's prog ID.
    Public ReadOnly Property ProgId() As String Implements SmartTagLib.ISmartTagRecognizer.ProgId
        Get
            Return "BookSmartTag.SmartTagRecognizer"
        End Get
    End Property

    ' Methods that describe the types of 
    ' smart tags the recognizer recognizes.

    ' The number of smart tag types this recognizer supports.
    ' This version only supports Office Smackdown.
    Public ReadOnly Property SmartTagCount() As Integer Implements SmartTagLib.ISmartTagRecognizer.SmartTagCount
        Get
            Return 1
        End Get
    End Property

    ' URI for the smart tag types that this
    ' rexognizer recognizes.
    Public ReadOnly Property SmartTagName(ByVal SmartTagID As Integer) As String Implements SmartTagLib.ISmartTagRecognizer.SmartTagName
        Get
            Return "http://www.vb-helper.com#BookSmartTag"
        End Get
    End Property

    ' URL where a user can download actions for this 
    ' smart tag if they are not already installed.
    Public ReadOnly Property SmartTagDownloadURL(ByVal SmartTagID As Integer) As String Implements SmartTagLib.ISmartTagRecognizer.SmartTagDownloadURL
        Get
            Return ""
        End Get
    End Property

    ' Methods that let the recognizer recognize.

    ' Make a list of terms we will recognize.
    Private m_Terms() As String = { _
            "book", _
            "oeuvre", _
            "tome", _
            "opus" _
        }

    ' Check for recognized text.
    Public Sub Recognize(ByVal examine_text As String, ByVal DataType As SmartTagLib.IF_TYPE, ByVal LocaleID As Integer, ByVal RecognizerSite As SmartTagLib.ISmartTagRecognizerSite) Implements SmartTagLib.ISmartTagRecognizer.Recognize
        Dim i As Integer
        Dim pos As Integer
        Dim property_bag As ISmartTagProperties

        ' Check each term.
        examine_text = examine_text.ToLower()
        For i = 0 To m_Terms.GetUpperBound(0)
            ' See if this term is present.
            pos = examine_text.IndexOf(m_Terms(i))
            Do While pos >= 0
                ' The term appears at position pos.
                ' Get a new property bag
                property_bag = RecognizerSite.GetNewPropertyBag()

                ' Commit the term. 
                ' Note that we add 1 to the position 
                ' because the property bag expects a 
                ' 1-based position.
                RecognizerSite.CommitSmartTag( _
                    "http://www.vb-helper.com#BookSmartTag", _
                    pos + 1, m_Terms(i).Length, property_bag)

                ' Look for the term's next occurrence.
                pos = examine_text.IndexOf(m_Terms(i), pos + m_Terms(i).Length)
            Loop
        Next i
    End Sub
End Class ' End SmartTagRecognizer.
