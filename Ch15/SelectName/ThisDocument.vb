' Other namespaces.
Imports MSForms = Microsoft.Vbe.Interop.Forms

Imports System.Windows.Forms
Imports Word = Microsoft.Office.Interop.Word

' Office integration attribute. Identifies the startup class for the document. Do not modify.
<Assembly: System.ComponentModel.DescriptionAttribute("OfficeStartupClass, Version=1.0, Class=SelectName.OfficeCodeBehind")>

Public Class OfficeCodeBehind

    Friend WithEvents ThisDocument As Word.Document
    Friend WithEvents ThisApplication As Word.Application

    ' Objects on the document.
    Private WithEvents cboName As MSForms.ComboBox

#Region "Generated initialization code"

    ' Default constructor.
    Public Sub New()
    End Sub

    ' Required procedure. Do not modify.
    Public Sub _Startup(ByVal application As Object, ByVal document As Object)
        ThisApplication = CType(application, Word.Application)
        ThisDocument = CType(document, Word.Document)

        If (ThisDocument.FormsDesign = True) Then
            ThisDocument.ToggleFormsDesign()
            ThisDocument_Open()
        End If
    End Sub

    ' Required procedure. Do not modify.
    Public Sub _Shutdown()
        ThisApplication = Nothing
        ThisDocument = Nothing
    End Sub

    ' Returns the control with the specified name in ThisDocument.
    Overloads Function FindControl(ByVal name As String) As Object
        Return FindControl(name, ThisDocument)
    End Function

    ' Returns the control with the specified name in the specified document.
    Overloads Function FindControl(ByVal name As String, ByVal document As Word.Document) As Object
        Try
            Dim inlineShape As Word.InlineShape
            For Each inlineShape In document.InlineShapes
                If (inlineShape.Type = Word.WdInlineShapeType.wdInlineShapeOLEControlObject) Then
                    Dim oleControl As Object = inlineShape.OLEFormat.Object
                    Dim oleControlType As Type = oleControl.GetType()
                    Dim oleControlName As String = CType(oleControlType.InvokeMember("Name", _
                    Reflection.BindingFlags.GetProperty, Nothing, oleControl, Nothing), String)
                    If (String.Compare(oleControlName, name, True, System.Globalization.CultureInfo.InvariantCulture) = 0) Then
                        Return oleControl
                    End If
                End If
            Next

            Dim shape As Word.Shape
            For Each shape In document.Shapes
                If (shape.Type = Microsoft.Office.Core.MsoShapeType.msoOLEControlObject) Then
                    Dim oleControl As Object = shape.OLEFormat.Object
                    Dim oleControlType As Type = oleControl.GetType()
                    Dim oleControlName As String = CType(oleControlType.InvokeMember("Name", _
                    Reflection.BindingFlags.GetProperty, Nothing, oleControl, Nothing), String)
                    If (String.Compare(oleControlName, name, True, System.Globalization.CultureInfo.InvariantCulture) = 0) Then
                        Return oleControl
                    End If
                End If
            Next

        Catch Ex As Exception
            ' Returns Nothing if the control is not found.
        End Try
        Return Nothing
    End Function
#End Region

    ' Called when the document is opened.
    Private Sub ThisDocument_Open() Handles ThisDocument.Open
        ' Find and initialize the ComboBox.
        cboName = FindControl("cboName")
        cboName.Clear()
        Try
            cboName.AddItem(ThisDocument.BuiltInDocumentProperties("Author").Value)
        Catch ex As Exception
        End Try
        cboName.AddItem("George Bush")
        cboName.AddItem("Annie Lennox")
        cboName.AddItem("Sergio Aragonés")
    End Sub

    ' Insert the selected name.
    Private Sub cboName_Click() Handles cboName.Click
        ' Find the bmName bookmark.
        Dim bm As Word.Bookmark
        Try
            bm = ThisDocument.Bookmarks("bmName")
        Catch ex As Exception
            MsgBox("Cannot find bookmark bmName")
            Exit Sub
        End Try

        ' Replace the bookmark's text with the selection.
        Dim rng As Word.Range = bm.Range
        rng.Text = cboName.Text

        ' Make the new text the bookmark.
        ThisDocument.Bookmarks.Add("bmName", rng)
    End Sub

    ' Called when the document is closed.
    Private Sub ThisDocument_Close() Handles ThisDocument.Close

    End Sub
End Class
