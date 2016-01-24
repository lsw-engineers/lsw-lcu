Attribute VB_Name = "WorksheetFixes"
Option Explicit

Function ApplyCellNames()

    Dim MaxCircuits As Byte
    Dim CurrentCkt As Byte
    Dim CurrentCell As Excel.Range
    Dim CurrentRow As Byte
    Dim CurrentCol As Byte
    Dim SchdPoles As Byte
    
    SchdPoles = GetPoles()
    
    Select Case GetInfo("SCHD_Type")
       
        Case "PANEL"
            If Range("I95").Value = "84" Then
                MaxCircuits = 84
                Else
                MaxCircuits = 42
            End If
            
            For CurrentCkt = 1 To MaxCircuits
            'Assign Cell Names for each Circuit
            
                CurrentRow = CurrentCkt + 11
                
                'Set Load Type (LT)
                If isEven(CurrentCkt) Then
                    Set CurrentCell = Cells(CurrentRow - 1, 12)
                Else
                    Set CurrentCell = Cells(CurrentRow, 2)
                End If
                CurrentCell.Name = "CKT_" & CurrentCkt & "_LT"
                
                'Set Breaker Poles (Poles)
                If isEven(CurrentCkt) Then
                    Set CurrentCell = Cells(CurrentRow, 10)
                Else
                    Set CurrentCell = Cells(CurrentRow + 1, 4)
                End If
                CurrentCell.Name = "CKT_" & CurrentCkt & "_Poles"
                
                'Set Volt-Amps (VA)
                CurrentCol = 6 + XLMod(WorksheetFunction.Round((CurrentCkt + 4) / 2, 0) / SchdPoles, 1) * 3
                Set CurrentCell = Cells(CurrentRow, CurrentCol)
                CurrentCell.Value = 0
                CurrentCell.Name = "CKT_" & CurrentCkt & "_VA"
                
                'Set LT Tables
                CurrentCol = 16 + XLMod(WorksheetFunction.Round((CurrentCkt + 4) / 2, 0) / SchdPoles, 1) * 3
                Set CurrentCell = Cells(CurrentRow, CurrentCol)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_C"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 3)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_L"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 6)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_M"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 9)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_ML"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 12)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_R"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 15)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_H"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 18)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_T"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 21)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_K"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 24)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_KQ"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 27)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_X"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 30)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_XL"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 33)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_XS"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 36)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_XQ"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 39)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_Z"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 42)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_ZL"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 45)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_ZS"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 48)
                CurrentCell.Name = "CKT_" & CurrentCkt & "_ZQ"
                Set CurrentCell = Cells(CurrentRow, CurrentCol + 51)
                
            Next
        
        Case "BUS"
            
    End Select

End Function
