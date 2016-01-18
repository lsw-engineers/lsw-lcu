VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents XL As Application
Attribute XL.VB_VarHelpID = -1

Private Sub XL_WorkbookActivate(ByVal Wb As Workbook)
'
'    If UCase(Wb.Name) = UCase("LCU.xla") Then Exit Sub
'
            
    If IsMember(Wb.Names, "LCU_Version") Then Call UnhideLCUMenu
    
    If (IsMember(Wb.Worksheets, "Panel") Or IsMember(Wb.Worksheets, "Calc")) And _
       Not IsMember(Wb.Names, "LCU_Version") Then Call UnhideOldLCUMenu
       
    'Range("BY10:CM10").Columns.Hidden = True
    'Range("A26:A36").Rows.Hidden = True
    'Excel.ActiveWindow.Zoom = 80
    'Range("A1").Select
    

End Sub

Private Sub XL_WorkbookDeactivate(ByVal Wb As Workbook)
'
'    If UCase(Wb.Name) = UCase("LCU.xla") Then Exit Sub
'
    If IsMember(Wb.Names, "LCU_Version") Then Call HideLCUMenu
    
    If (IsMember(Wb.Worksheets, "Panel") Or IsMember(Wb.Worksheets, "Calc")) And _
       Not IsMember(Wb.Names, "LCU_Version") Then Call HideOldLCUMenu

End Sub

Private Sub XL_WorkbookOpen(ByVal Wb As Workbook)
    
    Dim Sht As Excel.Worksheet
    Dim SheetName As Variant
    
    If IsMember(Wb.Names, "LCU_ADDIN") Then Exit Sub

    For Each Sht In Wb.Worksheets
    
        If Sht.ProtectContents = True Then _
            Sht.Protect UserInterfaceOnly:=True
    
    Next

    If IsMember(Wb.Worksheets, "Bus") And IsMember(Wb.Names, "LCU_Version") Then

'Temp provisions for PCC only --->
'
        Wb.Sheets("Bus").PageSetup.LeftFooter = "Last Revised &T &D"

        SheetName = Split(Wb.Name, " ")

        Wb.Sheets("Bus").PageSetup.RightFooter = SheetName(0)

        Wb.Sheets("Bus").Activate
        Wb.Sheets("Bus").Range("A1").Activate
        
        Call AutoHide

    End If

    If IsMember(Wb.Worksheets, "Panel") And IsMember(Wb.Names, "LCU_Version") Then

'Temp provisions for PCC only --->
'
        Wb.Sheets("Panel").PageSetup.LeftFooter = "Last Revised &T &D"

        SheetName = Split(Wb.Name, " ")

        Wb.Sheets("Panel").PageSetup.RightFooter = SheetName(0)

        Wb.Sheets("Panel").Activate
        Wb.Sheets("Panel").Range("A1").Activate

        Call SetCktDivisions

        Call AutoHide

    End If
 
 
 
            
End Sub
