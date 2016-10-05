VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' Remove custom buttons and menu items.
Private Sub Application_Quit()
    'MsgBox "Application_Quit"

    DeleteCommandbarItem "CustomTools", "SayHi button"
    DeleteCommandbarItem "Tools", "SayHi"

    ' If the CustomTools CommandBar is empty,
    ' delete it.
    On Error Resume Next
    If ActiveExplorer.CommandBars("CustomTools").Controls.Count = 0 Then
        ActiveExplorer.CommandBars("CustomTools").Delete
    End If
End Sub

' Install the Say Hi button and menu item.
Private Sub Application_Startup()
    'MsgBox "Application_Startup"

    ' Create the CustomTools CommandBar
    ' and ensure it is visible.
    MakeCommandBarVisible "CustomTools"

    ' Create the toolbar button.
    MakeCommandbarItem "CustomTools", "SayHi button", _
        msoButtonIcon, "SayHi", "Say Hi!", _
        "Execute the SayHi macro", _
        481

    ' Create a Tools menu item.
    MakeCommandbarItem "Tools", "SayHi", _
        msoButtonCaption, "SayHi", "Say Hi!", _
        "Execute the SayHi macro"
End Sub
