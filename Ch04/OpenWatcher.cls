VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Declare the variable to catch events.
Public WithEvents TheApp As Application
Attribute TheApp.VB_VarHelpID = -1

' A presentation is being opened.
Private Sub TheApp_PresentationOpen(ByVal Pres As Presentation)
    MsgBox "OpenWatcher: TheApp_PresentationOpen"

    ' Run the presentation's Open macro.
    Pres.Application.Run "Open"
End Sub
