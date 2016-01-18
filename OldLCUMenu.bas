Attribute VB_Name = "OldLCUMenu"
Option Explicit

Sub CreateOldLCUMenu()

    Call DeleteOldLCUMenu

    Dim HelpIndex As Long
    Dim OldLCUMenu As Object 'Needs a better dimension
    Dim MenuItem As Object 'Needs a better dimension
    Dim SubMenuItem As Object 'Needs a better dimension
    
    HelpIndex = CommandBars("Worksheet Menu Bar").Controls("Help").Index
    Set OldLCUMenu = CommandBars("Worksheet Menu Bar").Controls.Add _
        (Type:=msoControlPopup, _
         Before:=HelpIndex, _
         Temporary:=True)
    OldLCUMenu.Caption = "&Old LCU"
    
    Set MenuItem = OldLCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .Caption = "Convert &Panel..."
        .OnAction = "ConvertPanel"
    End With
    
    Set MenuItem = OldLCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .Caption = "Convert &Dist Calc..."
        .OnAction = "ConvertBus"
    End With

End Sub

Public Sub DeleteOldLCUMenu()
    On Error Resume Next
    CommandBars(1).Controls("Old LCU").Delete
End Sub
 
Public Sub UnhideOldLCUMenu()
    On Error Resume Next
    CommandBars(1).Controls("Old LCU").Visible = True
    
    If IsMember(ActiveWorkbook.Sheets, "Panel") Then
        CommandBars(1).Controls("Old LCU").Controls("Convert Panel...").Enabled = True
        CommandBars(1).Controls("Old LCU").Controls("Convert Dist Calc...").Enabled = False
    Else
        CommandBars(1).Controls("Old LCU").Controls("Convert Panel...").Enabled = False
        CommandBars(1).Controls("Old LCU").Controls("Convert Dist Calc...").Enabled = True
    End If
    
End Sub

Public Sub HideOldLCUMenu()
    On Error Resume Next
    CommandBars(1).Controls("Old LCU").Visible = False
End Sub

'Public Sub LoadOrUnhideOldLCUMenu()
'
'    If IsMember(CommandBars(1).Controls, "Old LCU") Then
'        UnhideOldLCUMenu
'    Else
'        CreateOldLCUMenu
'    End If
'
'End Sub



