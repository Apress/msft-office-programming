Attribute VB_Name = "OlCh04"

' Make a toolbar button or menu item.
Sub MakeCommandbarItem(ByVal commandbar_name As String, ByVal item_caption As String, ByVal button_style As MsoButtonStyle, ByVal on_action As String, Optional ByVal tooltip_text As String = "", Optional ByVal description_text As String = "", Optional ByVal face_id As Integer = 0)
Dim btn As CommandBarButton

    ' See if the button already exists.
    On Error Resume Next
    Set btn = ActiveExplorer.CommandBars(commandbar_name).Controls(item_caption)
    If Err.Number = 0 Then
        ' The button exists. Increment Tag.
        btn.Tag = CInt(btn.Tag) + 1
    Else
        ' The button doesn't exist. Create it.
        With ActiveExplorer.CommandBars(commandbar_name).Controls.Add( _
          Type:=msoControlButton)
            .Style = button_style
            .TooltipText = tooltip_text
            .OnAction = on_action
            .DescriptionText = description_text
            .FaceId = face_id
            .Caption = item_caption
            .Tag = "1"
        End With
    End If
End Sub
' Delete a toolbar button or menu item.
Sub DeleteCommandbarItem(ByVal commandbar_name As String, ByVal item_caption As String)
Dim btn As CommandBarButton

    ' Get a reference to the button.
    On Error Resume Next
    Set btn = ActiveExplorer.CommandBars(commandbar_name).Controls(item_caption)
    If Err.Number = 0 Then
        ' The button exists.
        ' Decrement its Tag property.
        btn.Tag = CInt(btn.Tag) - 1

        ' See if the count is zero.
        If CInt(btn.Tag) <= 0 Then
            ' No document needs this button.
            ' Delete it.
            ActiveExplorer.CommandBars(commandbar_name).Controls(item_caption).Delete
        End If
    End If
End Sub

' If the CommandBar doesn't exist, create it.
' Then ensure that the CommandBar is visible.
Sub MakeCommandBarVisible(ByVal commandbar_name As String)
Dim command_bar As CommandBar

    ' See if the CommandBar exists.
    On Error Resume Next
    Set command_bar = ActiveExplorer.CommandBars(commandbar_name)
    If Err.Number <> 0 Then
        ' The CommandBar doesn't exist. Create it.
        Set command_bar = ActiveExplorer.CommandBars.Add(Name:=commandbar_name)
        command_bar.Position = msoBarTop
    End If

    ' Make the CommandBar visible.
    command_bar.Visible = True
End Sub
' Delete this CommandBar.
Sub DeleteCommandBar(ByVal commandbar_name As String)
    On Error Resume Next
    ActiveExplorer.CommandBars(commandbar_name).Delete
End Sub
