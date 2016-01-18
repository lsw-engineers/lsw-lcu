Attribute VB_Name = "SpecialtyCalcs"
Option Explicit

Public Sub AddAENSCalc()

    Dim SchdSht As Excel.Worksheet
    Dim CopySht As Excel.Worksheet
    Dim CopyFrom As String
    Dim CopyTo As String
    
    Application.DisplayAlerts = False
    ScreenUpdates (False)
    InSub ("ON")
    
    Set SchdSht = ActiveWorkbook.Worksheets(GetSchdSht())
    Set CopySht = Workbooks("lcu.xla").Sheets("AENS")
    
    SchdSht.Activate
    
    If GetPoles() <> 3 Then
        MsgBox "AENS Load Management is only for 3-phase panelboards", vbInformation
        GoTo FinishUp
    End If
    
    Select Case GetInfo("SCHD_Type")
    Case "PANEL"
        CopyTo = "B68"
        CopyFrom = "B68:L78"
    Case "BUS"

    End Select
    
    CopySht.Range(CopyFrom).Copy SchdSht.Range(CopyTo)
    
    Call AutoHide
    
FinishUp:

    InSub ("OFF")
    ScreenUpdates (True)
    Application.DisplayAlerts = True

End Sub

