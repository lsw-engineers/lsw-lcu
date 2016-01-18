Attribute VB_Name = "Overloads"
Option Explicit

Public Sub IsOverLoaded(Sht As Excel.Worksheet)
    
    On Error Resume Next
    
    Dim SchdPoles As Byte
    Dim CurrentPole As Byte
    
    SchdPoles = GetPoles(Sht.Parent)
    
    '***** CHECK FOR OVERLOAD *****
    
    Select Case GetInfo("SCHD_Type", Sht.Parent.Name)
    
        Case "PANEL"
    
            For CurrentPole = 1 To SchdPoles
                If Sht.Cells((Sht.Range("Misc2_L1_VA").Row + 12), CurrentPole + 5).Value > Sht.Range("Mains_Amps").Value Then
                    Call SetHeaderColor(Sht, 3)
                    Exit Sub
                End If
            Next
            
            '***** CHECK FOR 80% WARNING *****
            For CurrentPole = 1 To SchdPoles
                If Sht.Cells((Sht.Range("Misc2_L1_VA").Row + 12), CurrentPole + 5).Value > (0.8 * Sht.Range("Mains_Amps").Value) Then
                    Call SetHeaderColor(Sht, 44)
                    Exit Sub
                End If
            Next
            
        Case "BUS"
        
            For CurrentPole = 1 To SchdPoles
                'If Sht.Cells( 46, CurrentPole + 3).Value > Sht.Range("Mains_Amps").Value Then
                If Sht.Cells((Sht.Range("Load1_LT").Row - 1) + 37, CurrentPole + 3).Value > Sht.Range("Mains_Amps").Value Then
                    Call SetHeaderColor(Sht, 3)
                    Exit Sub
                End If
            Next
            
            '***** CHECK FOR 80% WARNING *****
            For CurrentPole = 1 To SchdPoles
                If Sht.Cells((Sht.Range("Load1_LT").Row - 1) + 37, CurrentPole + 3).Value > (0.8 * Sht.Range("Mains_Amps").Value) Then
                    Call SetHeaderColor(Sht, 44)
                    Exit Sub
                End If
            Next
    
    End Select
    
    Call SetHeaderColor(Sht, Sht.Range("A1").Interior.ColorIndex)

End Sub


Private Sub SetHeaderColor(Sht As Excel.Worksheet, ByVal ColorIndexNo As Byte)
    
    Dim SchdPoles As Byte
    Dim Header As Excel.Range
        
    SchdPoles = GetPoles(Sht.Parent)
    
    Select Case GetInfo("SCHD_Type", Sht.Parent.Name)
    
        Case "PANEL"
    
            Set Header = Sht.Range("C8", Cells(11, 8 + SchdPoles))
            
        Case "BUS"
        
            Set Header = Sht.Range("C" & (Sht.Range("Load1_LT").Row - 4), Cells((Sht.Range("Load1_LT").Row - 2), 3 + SchdPoles))
            
    End Select
    
    If Header.Interior.ColorIndex <> ColorIndexNo Then
        Header.Interior.ColorIndex = ColorIndexNo
    End If
 
End Sub
