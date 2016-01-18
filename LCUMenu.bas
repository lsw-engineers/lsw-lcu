Attribute VB_Name = "LCUMenu"
Option Explicit

Sub CreateLCUMenu()

    Call DeleteLCUMenu

    Dim HelpIndex As Long
    Dim LCUMenu As Object 'Needs a better dimension
    Dim MenuItem As Object 'Needs a better dimension
    Dim SubMenuItem As Object 'Needs a better dimension
    
    HelpIndex = CommandBars("Worksheet Menu Bar").Controls("Help").Index
    Set LCUMenu = CommandBars("Worksheet Menu Bar").Controls.Add _
        (Type:=msoControlPopup, _
         Before:=HelpIndex, _
         Temporary:=True)
    LCUMenu.Caption = "&LCU"
    
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .FaceId = 2308
        .Caption = "&Connect/Link..."
        .OnAction = "AddConnectionDialog"
    End With
    
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .FaceId = 2309
        .Caption = "&Disconnect/Unlink..."
        .OnAction = "RemConnectionDialog"
    End With
    
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlPopup)
    With MenuItem
        .BeginGroup = True
        .Enabled = False
        .Caption = "NEC 220.21 &Noncoincident Loads"
    End With

        Set SubMenuItem = MenuItem.Controls.Add _
            (Type:=msoControlButton)
        With SubMenuItem
            .BeginGroup = True
            .Caption = "&Add..."
            .OnAction = "&NoncoincidentLoadsDialog"
        End With
        
        Set SubMenuItem = MenuItem.Controls.Add _
            (Type:=msoControlButton)
        With SubMenuItem
            .Caption = "&Remove"
            .OnAction = "&NoncoincidentExistingLoads"
        End With
    
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .Caption = "NEC 220.34 Optional Method - &Schools"
        .OnAction = "ToggleSchoolCalcs"
    End With
    
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlPopup)
    With MenuItem
        .Caption = "NEC 220.35 &Existing Loads"
    End With

        Set SubMenuItem = MenuItem.Controls.Add _
            (Type:=msoControlButton)
        With SubMenuItem
            .BeginGroup = True
            .Caption = "&Add..."
            .OnAction = "ExistingLoadsDialog"
        End With
        
        Set SubMenuItem = MenuItem.Controls.Add _
            (Type:=msoControlButton)
        With SubMenuItem
            .Caption = "&Remove"
            .OnAction = "RemoveExistingLoads"
        End With
        
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlPopup)
    With MenuItem
        .Caption = "Specialty Calcs"
    End With

        Set SubMenuItem = MenuItem.Controls.Add _
            (Type:=msoControlButton)
        With SubMenuItem
            .BeginGroup = True
            .Caption = "&Add AENS Load Management Calc"
            .OnAction = "AddAENSCalc"
        End With
 
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .FaceId = 1977
        .Caption = "&Update Circuit Divisions"
        .OnAction = "SetCktDivisions"
        .ShortcutText = "Alt+F5"
        .BeginGroup = True
    End With

    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .Caption = "&Toggle Color"
        .OnAction = "ToggleColor"
    End With

    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .Caption = "&Fix LoadType Formulas"
        .OnAction = "FixLTFormulas"
    End With
    
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .Caption = "&Reset All Loads"
        .OnAction = "ResetPanelLoads"
    End With
    
    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .Caption = "&Print All Schedules (This Project)"
        .OnAction = "PrintAllSchds"
    End With

    Set MenuItem = LCUMenu.Controls.Add _
        (Type:=msoControlButton)
    With MenuItem
        .Caption = "&About LCU..."
        .BeginGroup = True
        .OnAction = "About_LCU"
    End With

    If Application.UserName = "Adam J. Bagby" Then
    
        Set MenuItem = LCUMenu.Controls.Add _
            (Type:=msoControlPopup)
        With MenuItem
            .BeginGroup = True
            .Caption = "&Admin/Test"
        End With
        
            Set SubMenuItem = MenuItem.Controls.Add _
                (Type:=msoControlButton)
            With SubMenuItem
                .Caption = "Export all Names"
                .OnAction = "ExportAllNames"
            End With
            
            Set SubMenuItem = MenuItem.Controls.Add _
                (Type:=msoControlButton)
            With SubMenuItem
                .Caption = "Delete all Names"
                .OnAction = "DeleteAllNames"
            End With
            
            Set SubMenuItem = MenuItem.Controls.Add _
                (Type:=msoControlButton)
            With SubMenuItem
                .Caption = "Add All Names"
                .OnAction = "DefineAllNames"
            End With
            
            Set SubMenuItem = MenuItem.Controls.Add _
                (Type:=msoControlButton)
            With SubMenuItem
                .Caption = "Clean Up Names"
                .OnAction = "CleanUpNames"
            End With

            Set SubMenuItem = MenuItem.Controls.Add _
                (Type:=msoControlButton)
            With SubMenuItem
                .Caption = "Spanner..."
                .OnAction = "RunSpanner"
            End With
            Set SubMenuItem = MenuItem.Controls.Add _
                (Type:=msoControlButton)
            With SubMenuItem
                .Caption = "Clear Spanner Names"
                .OnAction = "DeleteSpannerNames"
            End With
            
    End If

End Sub

Public Sub DeleteLCUMenu()
    On Error Resume Next
    CommandBars(1).Controls("LCU").Delete
End Sub
 
Public Sub UnhideLCUMenu()
    On Error Resume Next
    CommandBars(1).Controls("LCU").Visible = True
    
    If GetInfo("SCHD_Type") = "PANEL" Then
        CommandBars(1).Controls("LCU").Controls("Update Circuit Divisions").Enabled = True
    Else
        CommandBars(1).Controls("LCU").Controls("Update Circuit Divisions").Enabled = False
    End If
    
End Sub

Public Sub HideLCUMenu()
    On Error Resume Next
    CommandBars(1).Controls("LCU").Visible = False
End Sub

'Public Sub LoadOrUnhideLCUMenu()
'
'    If IsMember(CommandBars(1).Controls, "LCU") Then
'        UnhideLCUMenu
'    Else
'        CreateLCUMenu
'    End If
'
'End Sub

