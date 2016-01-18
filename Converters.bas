Attribute VB_Name = "Converters"
Option Explicit

Private Function parseforfilename(Formula As String) As String
    
    Dim FileName As String
    Dim Path As String
    
    FileName = Formula
    
    FileName = Right(FileName, (Len(FileName) - InStr(1, FileName, "[")))
    
    FileName = Left(FileName, (InStr(1, FileName, "]") - 1))
    
    Path = Formula
    
    Path = Right(Path, (Len(Path) - 2))
    
    Path = Left(Path, (InStr(1, Path, "[") - 1))
    
    parseforfilename = UCase(Path & FileName)

End Function

Public Sub ConvertPanel()

    Dim OldShtName As String
    Dim OldSht As Excel.Worksheet
    Dim NewSht As Excel.Worksheet
    Dim ShtToInsert As Excel.Worksheet
    Dim CurrentCkt As Byte
    Dim TestRange As Excel.Range
    Dim FileName As String
    Dim CircuitNo As Variant
    Dim CktsToCheck As Variant
    Dim Counter As Byte
    Dim CktNo As Byte
       
    PleaseWait.Show 0
    PleaseWait.Repaint

    ScreenUpdates (False)
    
    DeleteAllNames (True)

    Set OldSht = ActiveSheet
    Set ShtToInsert = Workbooks("lcu.xla").Sheets("Panel")
    
    OldSht.Name = "Old " & OldSht.Name
    
    FileName = OldSht.Parent.Path & "Temp_mod.bas"
    
    ShtToInsert.Parent.VBProject.VBComponents("LocalFunctions").Export FileName
    
    OldSht.Parent.VBProject.VBComponents.Import FileName
    
    Kill FileName
    
    ShtToInsert.Copy OldSht

    Set NewSht = ActiveSheet
    
    NewSht.Protect UserInterfaceOnly:=True
    
    ' ***** SETUP HIDDEN ROWS *****
    
    For Counter = 6 To 47

        If OldSht.Rows(Counter).Hidden = True Then _
            NewSht.Rows(Counter + 6).Hidden = True

    Next

    ' ***** XFER CKT INFO *****

    Dim XferRanges As Variant

    XferRanges = Array(Array("B6:C47", "B12:C53") _
                     , Array("E6:E47", "D12:D53") _
                     , Array("G6:I47", "F12:H53") _
                     , Array("K6:M47", "J12:L53") _
                     , Array("C48", "C54") _
                     , Array("G48", "F54") _
                     , Array("H48", "G54") _
                     , Array("I48", "H54") _
                       )
    
    For Counter = 0 To UBound(XferRanges, 1)
        OldSht.Range(XferRanges(Counter)(0)).Copy
        
        If NewSht.Range(XferRanges(Counter)(1)).MergeCells = True Then
            NewSht.Range(XferRanges(Counter)(1)).Value = _
                    OldSht.Range(XferRanges(Counter)(0)).Text
        Else
            NewSht.Range(XferRanges(Counter)(1)).PasteSpecial xlPasteValues
        End If
    
    Next


    For Counter = 6 To 47
    
        Set TestRange = OldSht.Range("D" & Counter)
    
        If TestRange.Text <> "" Then
        
            NewSht.Range("C" & (6 + Counter)).Value = _
                TestRange.Offset(0, -1).Value & " " & _
                TestRange.Value
        
        End If
    
    Next
    
    ' ***** COPY NOTES *****
    
    OldSht.Range("J50:L56").Copy
    NewSht.Range("I58").PasteSpecial xlPasteValues
    
    For Counter = 58 To 60
        If NewSht.Range("I" & Counter).Text = "" Then _
            NewSht.Range("I" & Counter).Value = " "
        
    Next

    ' ***** XFER HEADER INFO *****
    
    NewSht.Range("D2").Value = OldSht.Range("D2").Value ' Name
    NewSht.Range("D4").Value = OldSht.Range("D3").Value ' Mounting
    NewSht.Range("K2").Value = OldSht.Range("H3").Value ' Type
    NewSht.Range("K5").Value = OldSht.Range("L3").Value ' AIC
    
    ' Mains Amps
    
    If IsNumeric(Left(OldSht.Range("H2").Value, 3)) Then
    
        NewSht.Range("D5").Value = Left(OldSht.Range("H2").Value, 3)
    
    ElseIf IsNumeric(Left(OldSht.Range("H2").Value, 2)) Then
        
        NewSht.Range("D5").Value = Left(OldSht.Range("H2").Value, 2)
    
    End If
    
    ' Mains Type
        
    If InStr(1, OldSht.Range("H2").Value, "L", vbTextCompare) Then _
        NewSht.Range("D6").Value = "Lugs Only"
        
    If InStr(1, OldSht.Range("H2").Value, "C", vbTextCompare) Then _
        NewSht.Range("D6").Value = "Circuit Breaker"
        
    ' Voltage
    
    Select Case Left(OldSht.Range("L2").Value, 4)
        Case "120/", "208/", "208Y"
            NewSht.Range("K4").Value = "208Y/120V, 3ø, 4W"
        Case "208V"
            NewSht.Range("K4").Value = "208V, 3ø, 3W"
        Case "277/", "480/", "480Y"
            NewSht.Range("K4").Value = "480Y/277V, 3ø, 4W"
        Case "480V"
            NewSht.Range("K4").Value = "480V, 3ø, 3W"
    End Select
    
    
    ' ***** CONVERT 'R' to 'H' for HOSPITAL PNLS *****
    If OldSht.Range("N1").Value = "HOSPITAL" Then
        For CurrentCkt = 1 To 42
            Set TestRange = NewSht.Range("CKT_" & CurrentCkt & "_LT")
            If TestRange.Value = "R" Then TestRange.Value = "H"
        Next
    End If
        
        
    ' ***** DEAL WITH LINKED SCHEDULES *****

    CircuitNo = Array(1, 2, 7, 8, 13, 14, 19, 20, 25, 26, 31, 32, 37, 38)
    CktsToCheck = Array("G6", "G7", "G12", "G13", "G18", "G19", "G24", "G25", _
                        "G30", "G31", "G36", "G37", "G42", "G43")
    
    For Counter = 1 To UBound(CktsToCheck, 1)
    
        Set TestRange = OldSht.Range(CktsToCheck(Counter))
        
        If Left(TestRange.Formula, 2) = "='" Then
            
            FileName = parseforfilename(TestRange.Formula)
            
            CktNo = CircuitNo(Counter)
            
            If GetInfo("LCU_Version", FileName) <> "INVALID" Then
                
                Call LinkSchedule(FileName, "CKT", 3, CktNo)
            
            Else
            
                MsgBox "Linking could not be restored to the following load schedule:" & _
                        vbCrLf & vbCrLf & FileName & vbCrLf & vbCrLf & _
                        "The load schedule being attached is of an incompatable version.  Please " & _
                        "convert the referenced" & vbCrLf & "load schedule to the current version " & _
                        "and then manually reconnect/relink the file", vbCritical, _
                        "Error Linking Schedules"
                        
                NewSht.Range("CKT_" & CktNo & "_VA").Value = "Broken Link to " & FileNameOnly(FileName)
                NewSht.Range("CKT_" & (CktNo + 2) & "_VA").Value = "Broken Link to " & FileNameOnly(FileName)
                NewSht.Range("CKT_" & (CktNo + 4) & "_VA").Value = "Broken Link to " & FileNameOnly(FileName)
                
            End If
            
        End If
    
    Next
    
    Set TestRange = OldSht.Range("G48")
        
    If Left(TestRange.Formula, 2) = "='" Then
        
        FileName = parseforfilename(TestRange.Formula)
        
        If GetInfo("LCU_Version", FileName) <> "INVALID" Then
            
            CktNo = 1
            
            Call LinkSchedule(FileName, "Misc1", 3, CktNo)
        
        Else
        
            MsgBox "Linking could not be restored to the following load schedule:" & _
                    vbCrLf & vbCrLf & FileName & vbCrLf & vbCrLf & _
                    "The load schedule being attached is of an incompatable version.  Please " & _
                    "convert the referenced" & vbCrLf & "load schedule to the current version " & _
                    "and then manually reconnect/relink the file", vbCritical, _
                    "Error Linking Schedules"
                    
            NewSht.Range("F54:H54").Value = "Broken Link to " & FileNameOnly(FileName)
            
        End If
        
    End If
    
    NewSht.Range("A1").Activate
    
    NewSht.Calculate
    
    Call SetCktDivisions
    Call AutoHide
    
    ScreenUpdates (True)
    
    PleaseWait.Hide
    Unload PleaseWait
        
    MsgBox "Panel has been converted.  Please check for errors and then disgard the " _
            & OldSht.Name & " worksheet.", vbInformation

End Sub


Public Sub ConvertBus()

    Dim OldShtName As String
    Dim OldSht As Excel.Worksheet
    Dim NewSht As Excel.Worksheet
    Dim ShtToInsert As Excel.Worksheet
    Dim CurrentCkt As Byte
    Dim TestRange As Excel.Range
    Dim FileName As String
    Dim Counter As Byte
    Dim AmpsVolts As Variant
       
    PleaseWait.Show 0
    PleaseWait.Repaint

    ScreenUpdates (False)
    
    DeleteAllNames (True)

    Set OldSht = ActiveSheet
    Set ShtToInsert = Workbooks("lcu.xla").Sheets("Bus")
    
    OldSht.Name = "Old " & OldSht.Name
    
    FileName = OldSht.Parent.Path & "Temp_mod.bas"
    
    ShtToInsert.Parent.VBProject.VBComponents("LocalFunctions").Export FileName
    
    OldSht.Parent.VBProject.VBComponents.Import FileName
    
    Kill FileName
    
    ShtToInsert.Copy OldSht

    Set NewSht = ActiveSheet
    
    NewSht.Protect UserInterfaceOnly:=True
    
    ' ***** SETUP HIDDEN ROWS *****
    
    For Counter = 8 To 32

        If OldSht.Rows(Counter).Hidden = True Then _
            NewSht.Rows(Counter + 2).Hidden = True

    Next

    ' ***** XFER LOAD INFO *****
    
    OldSht.Range("B8:F32").Copy
    NewSht.Range("B10").PasteSpecial xlPasteValues

    ' ***** XFER HEADER INFO *****
    
    OldSht.Range("C2").Copy
    NewSht.Range("C5").PasteSpecial xlPasteValues
    
    ' Mains Amps
    
    NewSht.Range("Mains_Amps").Value = GetAmpsVolts(OldSht.Range("C4").Value)(0)
    
    ' Voltage
    
    Select Case Left(GetAmpsVolts(OldSht.Range("C4").Value)(1), 4)
        Case "120/", "208/", "208Y"
            NewSht.Range("Voltage").Value = "208Y/120V, 3ø, 4W"
        Case "208V"
            NewSht.Range("Voltage").Value = "208V, 3ø, 3W"
        Case "277/", "480/", "480Y"
            NewSht.Range("Voltage").Value = "480Y/277V, 3ø, 4W"
        Case "480V"
            NewSht.Range("Voltage").Value = "480V, 3ø, 3W"
        Case ""
            NewSht.Range("Voltage").Value = ""
    End Select
    
    ' ***** DEAL WITH LINKED SCHEDULES *****

    For Counter = 8 To 32

        Set TestRange = OldSht.Range("D" & Counter)

        If Left(TestRange.Formula, 2) = "='" Then

            FileName = parseforfilename(TestRange.Formula)

            If GetInfo("LCU_Version", FileName) <> "INVALID" Then

                Call LinkSchedule(FileName, "Load" & (Counter - 7), 3, 1)

            Else

                MsgBox "Linking could not be restored to the following load schedule:" & _
                        vbCrLf & vbCrLf & FileName & vbCrLf & vbCrLf & _
                        "The load schedule being attached is of an incompatable version.  Please " & _
                        "convert the referenced" & vbCrLf & "load schedule to the current version " & _
                        "and then manually reconnect/relink the file", vbCritical, _
                        "Error Linking Schedules"

                NewSht.Range("Load" & (Counter - 7) & "_L1_VA").Value = "Broken Link to " & FileNameOnly(FileName)
                NewSht.Range("Load" & (Counter - 7) & "_L2_VA").Value = "Broken Link to " & FileNameOnly(FileName)
                NewSht.Range("Load" & (Counter - 7) & "_L3_VA").Value = "Broken Link to " & FileNameOnly(FileName)

            End If

        End If

    Next
    
    NewSht.Range("A1").Activate
    
    NewSht.Calculate

    Call AutoHide
    
    ScreenUpdates (True)
    
    PleaseWait.Hide
    Unload PleaseWait
        
    MsgBox "Distribution calc has been converted.  Please check for errors and then disgard the " _
            & OldSht.Name & " worksheet.", vbInformation

End Sub

Private Function GetAmpsVolts(Text As String) As Variant

    Dim AmpsVolts As Variant
    Dim Amps As String
    Dim Volts As String
    
    If InStr(1, Text, "-") = 0 Then
        GetAmpsVolts = Array("", "")
        Exit Function
    End If
    
    AmpsVolts = Split(Text, "-")
    
    Amps = Trim(AmpsVolts(0))
    
    Do Until IsNumeric(Right(Amps, 1))
    
        If Amps <> "" Then
            Amps = Left(Amps, Len(Amps) - 1)
        Else
            Amps = 0
        End If
    
    Loop
    
    Volts = Trim(AmpsVolts(1))
    
'    MsgBox "Amps: " & Amps & vbCrLf & _
'           "Volts: " & Volts

    GetAmpsVolts = Array(Amps, Volts)
    
End Function

