Attribute VB_Name = "NEC220_35_ExistingLoads"
Option Explicit

Public Sub ExistingLoadsDialog()

    If HasExLoads() Then
        ActiveWorkbook.Worksheets("Existing Loads").Activate
        
        If MsgBox("Existing Loads already setup." & vbCrLf & vbCrLf & _
               "Do you wish to replace the current Existing Loads?", _
               vbYesNo + vbExclamation, "Replace Existing Loads") = vbYes Then
        
            Call RemoveExistingLoads
        
        Else
            
            Exit Sub
        
        End If
        
    End If

    ExistingLoadsForm.Show

End Sub

Public Sub AddExistingLoads(InputUnits As String, InputMethod As String)

    'InputUnits shall be "Amps", "kVA", or "kW"
    'InputMethod shall be "Total" or "Individual"
    
    Dim SchdSht As Excel.Worksheet
    Dim ShtToInsert As Excel.Worksheet
    Dim ExLoadSht As Excel.Worksheet
    Dim Cell As Excel.Range
    Dim UpLeftCell As Excel.Range
    Dim SchdPoles As Byte
    Dim ExLoadRow As Byte
    Dim ExLoadCol As Byte
    Dim PhaseNameRow As Byte
    
    Set SchdSht = ActiveWorkbook.Worksheets(1)
    
    ScreenUpdates (False)
    InSub ("ON")
    
    Set ShtToInsert = Workbooks("lcu.xla").Sheets("Existing Loads")
    
    ShtToInsert.Copy After:=SchdSht
    
    Set ExLoadSht = ActiveSheet

    Set UpLeftCell = ExLoadSht.Range("B2")
    
    SchdPoles = GetPoles()
    
    Select Case GetInfo("SCHD_Type")
    
    Case "PANEL"
    
        ExLoadRow = 56
        ExLoadCol = 5
        PhaseNameRow = 11
        
        SchdSht.Cells(ExLoadRow, 3).Value = "Existing Load per NEC 220.35"
        SchdSht.Cells(ExLoadRow, 6 + SchdPoles).Value = "(Maximum Demand at 125%)"
        
    Case "BUS"
        
        ExLoadRow = 9
        ExLoadCol = 3
        PhaseNameRow = 8
    
    End Select
    
    UpLeftCell.Offset(2, 0).Value = "Maximum Demand in " & InputUnits & " (" & InputMethod & "):"
    
    Select Case InputMethod
    
    Case "Total" '****************************************
    
        With UpLeftCell.Offset(2, 1) 'Setup Input Line
            .HorizontalAlignment = xlCenter
            .BorderAround (xlContinuous)
            .Interior.ColorIndex = xlColorIndexNone
        End With
    
        If InputUnits = "kW" Then 'Setup Power Factor Line
        
            UpLeftCell.Offset(3, 0).Value = "Assumed or Measured Power Factor:"
        
            With UpLeftCell.Offset(3, 1)
                .Value = 0.8
                .HorizontalAlignment = xlCenter
                .BorderAround (xlContinuous)
                .Interior.ColorIndex = xlColorIndexNone
            End With
                    
        End If
        
        Dim CurrentPole As Byte
        
        For CurrentPole = 1 To SchdPoles 'Setup Formulas in Schedule
        
            With SchdSht.Cells(ExLoadRow, CurrentPole + ExLoadCol)
                .Interior.ColorIndex = 8
                
                Select Case InputUnits
                Case "kW"
                    .Formula = "=ROUND(1.25*(((" & UpLeftCell.Offset(2, 1). _
                                Address(External:=True) & "/" & UpLeftCell.Offset(3, 1). _
                                Address(External:=True) & ")*1000)/" & SchdPoles & "),0)"
                Case "kVA"
                    .Formula = "=ROUND(1.25*(" & UpLeftCell.Offset(2, 1). _
                                Address(External:=True) & "/" & SchdPoles & ")*1000,0)"
                Case "Amps"
                    .Formula = "=ROUND(1.25*" & UpLeftCell.Offset(2, 1). _
                                Address(External:=True) & "*Voltage_LN,0)"
                End Select
                
            End With
        Next
    
    Case "Individual" '*******************************
    
        For CurrentPole = 1 To SchdPoles 'Setup Phase Description Line
        
            With UpLeftCell.Offset(1, CurrentPole)
                .Value = SchdSht.Cells(PhaseNameRow, CurrentPole + ExLoadCol).Value
                .HorizontalAlignment = xlCenter
            End With
        Next
         
        For CurrentPole = 1 To SchdPoles 'Setup Input Line
        
            With UpLeftCell.Offset(2, CurrentPole)
                .HorizontalAlignment = xlCenter
                .BorderAround (xlContinuous)
                .Interior.ColorIndex = xlColorIndexNone
            End With
        Next
        
        If InputUnits = "kW" Then 'Setup Power Factor Line
        
            UpLeftCell.Offset(3, 0).Value = "Assumed or Measured Power Factor:"
        
            For CurrentPole = 1 To SchdPoles
            
                With UpLeftCell.Offset(3, CurrentPole)
                    .Value = 0.8
                    .HorizontalAlignment = xlCenter
                    .BorderAround (xlContinuous)
                    .Interior.ColorIndex = xlColorIndexNone
                End With
            Next
                    
        End If
        
        For CurrentPole = 1 To SchdPoles 'Setup Formulas in Schedule
        
            With SchdSht.Cells(ExLoadRow, CurrentPole + ExLoadCol)
                .Interior.ColorIndex = 8
                
                Select Case InputUnits
                Case "kW"
                    .Formula = "=ROUND(1.25*(" & UpLeftCell.Offset(2, CurrentPole). _
                                Address(External:=True) & "/" & UpLeftCell. _
                                Offset(3, CurrentPole).Address(External:=True) _
                                & ")*1000,0)"
                Case "kVA"
                    .Formula = "=ROUND(1.25*" & UpLeftCell.Offset(2, CurrentPole). _
                                Address(External:=True) & "*1000,0)"
                Case "Amps"
                    .Formula = "=ROUND(1.25*" & UpLeftCell.Offset(2, CurrentPole). _
                                Address(External:=True) & "*Voltage_LN,0)"
                End Select
                
            End With
        Next
        
    End Select

FinishUp:

    InSub ("OFF")
    ScreenUpdates (True)

    SchdSht.Rows(ExLoadRow).Hidden = False
    
    ExLoadSht.Range("A1").Activate
    
End Sub

Public Sub RemoveExistingLoads()

    Dim SchdPoles As Byte

    ScreenUpdates (False)
    Application.DisplayAlerts = False
    
    If HasExLoads() Then
        ActiveWorkbook.Sheets("Existing Loads").Delete
        ActiveWorkbook.Sheets(GetSchdSht()).Activate
    End If

    SchdPoles = GetPoles()
    
    Select Case GetInfo("SCHD_Type")
    
        Case "PANEL"
        
            With Range(Cells(56, 3).Address(False, False) & ":" _
                       & Cells(56, 8 + SchdPoles).Address(False, False))
                .ClearContents
                .Interior.ColorIndex = 15
            End With
            
            Range("A56").EntireRow.Hidden = True
            
        Case "BUS"
        
            With Range(Cells(9, 4).Address(False, False) & ":" _
                       & Cells(9, 3 + SchdPoles).Address(False, False))
                .ClearContents
                .Interior.ColorIndex = 15
            End With
            
            Range("A9").EntireRow.Hidden = True
        
    End Select

    Application.DisplayAlerts = True
    ScreenUpdates (True)

End Sub

Private Function HasExLoads() As Boolean

    If IsMember(ActiveWorkbook.Worksheets, "Existing Loads") Then
        HasExLoads = True
    Else
        HasExLoads = False
    End If
    
End Function

