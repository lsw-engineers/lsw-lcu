Attribute VB_Name = "Disconnecting"
Option Explicit

Public Sub RemConnectionDialog()

    RemConnectionForm.Show

End Sub

Public Sub FixLTFormulas()

    Dim CurrentCkt As Byte
    Dim CurrentLine As Byte
    Dim TestFormula As String
    Dim SchdPoles As Byte
    
    PleaseWait.Show 0
    PleaseWait.Repaint
    
    ActiveWorkbook.Sheets(GetSchdSht()).Activate
    
    InSub (True)
    ScreenUpdates (False)
    
    SchdPoles = GetPoles()
    
    Select Case GetInfo("SCHD_Type")
    
        Case "PANEL"
        
            For CurrentCkt = 1 To 42
            
                TestFormula = Range("CKT_" & CurrentCkt & "_VA").Formula
                
                If Left(TestFormula, 2) <> "='" Then Call ResetLoadFormulas("CKT", 1, CurrentCkt, True)
            
            Next
            
            TestFormula = Range("Misc1_L1_VA").Formula
            If Left(TestFormula, 2) <> "='" Then Call ResetLoadFormulas("Misc1", SchdPoles, 1, True)
            
            TestFormula = Range("Misc2_L1_VA").Formula
            If Left(TestFormula, 2) <> "='" Then Call ResetLoadFormulas("Misc2", SchdPoles, 1, True)
            
        Case "BUS"
        
            For CurrentLine = 1 To SchdPoles
            
                For CurrentCkt = 1 To 25
                
                    TestFormula = Range("Load" & CurrentCkt & "_L" & CurrentLine & "_VA").Formula
                    If Left(TestFormula, 2) <> "='" Then Call ResetLoadFormulas("Load" & CurrentCkt, 1, CurrentLine, True)
            
                Next
                
                TestFormula = Range("Misc1_L" & CurrentLine & "_VA").Formula
                If Left(TestFormula, 2) <> "='" Then Call ResetLoadFormulas("Misc1", 1, CurrentLine, True)
                
            Next

    
    End Select
    
    ScreenUpdates (True)
    InSub (False)
        
    PleaseWait.Hide
    Unload PleaseWait
        
    MsgBox "All loads have been left as is and all load type formulas have been restored.", _
            vbInformation

End Sub


Public Sub ResetPanelLoads()

    Dim Response As Variant
    Dim i As Byte
    Dim SchdPoles As Byte
    
    ActiveWorkbook.Sheets(GetSchdSht()).Activate
    
    Response = MsgBox("Are you sure you want to reset all VA values and loadtype formulas?" _
                        & vbCrLf & "(Any linked schedules will be unlinked by this process.)", _
                        vbExclamation + vbYesNo, "Reset All Loads")
                        
    If Response <> vbYes Then Exit Sub
        
    PleaseWait.Show 0
    PleaseWait.Repaint
    
    InSub (True)
    ScreenUpdates (False)
    
    SchdPoles = GetPoles()
    
    Select Case GetInfo("SCHD_Type")
    
    Case "PANEL"
        
        For i = 1 To 42
            Call ResetLoadFormulas("CKT", 1, i)
        Next
        
        Call ResetLoadFormulas("Misc1", SchdPoles, 1)
        Call ResetLoadFormulas("Misc2", SchdPoles, 1)
        
    Case "BUS"
    
        For i = 1 To 25
            Call ResetLoadFormulas("Load" & i, SchdPoles, 1)
        Next
        
        Call ResetLoadFormulas("Misc1", SchdPoles, 1)
    
    End Select
    
    ScreenUpdates (True)
    InSub (False)
        
    PleaseWait.Hide
    Unload PleaseWait
        
    MsgBox "All loads have been set to 0 and all load type formulas have been restored."
        
End Sub


Public Sub ResetLoadFormulas(ConnType As String, NoPoles As Byte, _
                             FirstCktNo As Byte, Optional IgnoreVA As Boolean = False)
    
    'Acceptable ConnType values are "CKT" "Misc1" "Misc2"
    'FirstCktNo will be an integer 1 to 42
    'NoPoles will be an integer 1 to 3
    
    Dim SchdPoles As Byte
    Dim CellToModify As String
    Dim UseFormula As String
    Dim AssocCkts As Variant
    Dim CellTypes As Variant
    Dim CurrentCkt As Variant
    Dim CurrentType As Variant
    Dim LTCellName As String
    Dim Row As Long
    Dim SchdType As String
            
    SchdPoles = GetPoles()
    SchdType = GetInfo("SCHD_Type")
    
    If NoPoles > SchdPoles Then
        MsgBox "Error.  Tried to disconnect too many poles."
        ScreenUpdates (True)
        Exit Sub
    End If

    Select Case NoPoles
    
    Case "1" And ConnType = "CKT"
        AssocCkts = Array(FirstCktNo)
    Case "2" And ConnType = "CKT"
        AssocCkts = Array(FirstCktNo, FirstCktNo + 2)
    Case "3" And ConnType = "CKT"
        AssocCkts = Array(FirstCktNo, FirstCktNo + 2, FirstCktNo + 4)

    Case "1" And ConnType <> "CKT"
        AssocCkts = Array("L1")
    Case "2" And ConnType <> "CKT"
        AssocCkts = Array("L1", "L2")
    Case "3" And ConnType <> "CKT"
        AssocCkts = Array("L1", "L2", "L3")
        
    End Select
    
    CellTypes = Array("VA", "LT", "C", "L", "M", "ML", "R", "H", "T", "K", "KQ", _
                            "X", "XL", "XS", "XQ", "Z", "ZL", "ZS", "ZQ")
    
    For Each CurrentCkt In AssocCkts
    
        For Each CurrentType In CellTypes
        
            If ConnType <> "CKT" And CurrentType = "LT" Then
            
                CellToModify = ConnType & "_" & CurrentType
            
            Else
                
                CellToModify = ConnType & "_" & CurrentCkt & "_" & CurrentType
            
            End If

            If ConnType <> "CKT" Then
                
                LTCellName = ConnType & "_LT"
            
            Else
            
                LTCellName = ConnType & "_" & CurrentCkt & "_LT"
                
            End If


            Select Case CurrentType
            
                Case "VA"
                UseFormula = ""
                
                Case "LT"
                UseFormula = ""
                
                Case "C", "L", "M", "R", "H", "T", "K", "X", "Z"
                UseFormula = "=IF(ISERR(SEARCH(""" & CurrentType & """," & LTCellName & ")),0," & ConnType & "_" & CurrentCkt & "_VA)"
                
                Case "ML", "XL", "XS", "ZL", "ZS"
                UseFormula = "=" & ConnType & "_" & CurrentCkt & "_" & Left(CurrentType, 1)
                
                Case "KQ"
                
                    If SchdType = "PANEL" Then
                        
                        UseFormula = "=IF(" & ConnType & "_" & CurrentCkt & "_" & Left(CurrentType, 1) & "=0,0,IF(" & ConnType & "_" & CurrentCkt & "_Poles>0,1,0))"
                    
                    ElseIf SchdType = "BUS" Then
                    
                        If CurrentCkt = "L1" Then
                            
                            UseFormula = "=IF(Sum(" & Range(ConnType & "_" & CurrentCkt & "_" & Left(CurrentType, 1), Range(ConnType & "_" & CurrentCkt & "_" & Left(CurrentType, 1)).Offset(0, SchdPoles - 1)).Address(False, False) & ")>0,1,0)"
                        
                        Else
                            
                            UseFormula = ""
                            
                        End If
                    
                    End If
                            
                Case "XQ", "ZQ"
                UseFormula = "=IF(" & ConnType & "_" & CurrentCkt & "_" & Left(CurrentType, 1) & ">0,1,0)"

            End Select
            
            If (CurrentType = "VA" Or CurrentType = "LT") And IgnoreVA = True Then
                
                'Do Nothing
                
            Else
                
                Range(CellToModify).Formula = UseFormula
                
            End If
            
        Next
    
    Next
    
    If ConnType <> "CKT" And SchdType = "PANEL" Then
    
        Row = Range(ConnType & "_L1_VA").Row
        
        Cells(Row, 6 + SchdPoles).ClearContents
    
    End If

End Sub


