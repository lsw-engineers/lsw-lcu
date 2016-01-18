Attribute VB_Name = "NEC220_34_SchoolCalcs"
Option Explicit

Public Sub ToggleSchoolCalcs()

    Dim SchdSht As Excel.Worksheet
    Dim CopySht As Excel.Worksheet
    Dim CopyFrom As String
    Dim CopyTo As String
    Dim MasterSht As String
    
    Application.DisplayAlerts = False
    ScreenUpdates (False)
    InSub ("ON")
    
    Set SchdSht = ActiveWorkbook.Worksheets(GetSchdSht())
    Set CopySht = Workbooks("lcu.xla").Sheets("Noncoincident Loads")
    
    SchdSht.Activate
    
    If GetPoles() <> 3 Then
        MsgBox "School calcs are not setup for single phase schedules", vbInformation
        GoTo FinishUp
    End If
    
    Select Case GetInfo("SCHD_Type")
    Case "PANEL"
        CopyTo = "C58"
        CopyFrom = "C58:H68"
        MasterSht = "Panel"
    Case "BUS"
        CopyTo = "C37"
        CopyFrom = "C37:F48"
        MasterSht = "Bus"
    End Select
        
    If Left(SchdSht.Range(CopyTo).Value, 4) = "Area" Then
        MsgBox "School Calcs Currently Enabled... Switching to Normal Calcs", vbInformation
        Set CopySht = Workbooks("lcu.xla").Sheets(MasterSht)
    Else
        MsgBox "Normal Calcs Currently Enabled... Switching to School Calcs", vbInformation
        Set CopySht = Workbooks("lcu.xla").Sheets("School Calc")
    End If
    
    CopySht.Range(CopyFrom).Copy SchdSht.Range(CopyTo)
    
    Call AutoHide
    
FinishUp:

    InSub ("OFF")
    ScreenUpdates (True)
    Application.DisplayAlerts = True

End Sub
