Attribute VB_Name = "AutoExec"
' An instance of this object to watch events.
Public g_ApplicationEventWatcher As ApplicationEventWatcher

' Create the ApplicationEventWatcher instance.
Sub Main()
    Set g_ApplicationEventWatcher = New ApplicationEventWatcher
End Sub
