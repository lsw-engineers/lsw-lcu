Attribute VB_Name = "Asthetics"
Option Explicit

Public Sub SetCktDivisions()

    Application.EnableEvents = False
    ScreenUpdates (False)

    Dim CurrentCktNo As Byte
    Dim CurrentCkt As Excel.Range
    Dim NoPoles As String
    Dim SchdPoles As Byte
    Dim MaxCircuits As Byte
    
    SchdPoles = GetPoles()
    
    If Range("I95").Value = "84" Then
        MaxCircuits = 84
        Else
        MaxCircuits = 42
    End If
     
    For CurrentCktNo = 1 To (MaxCircuits - 1) Step 2 'Cycle Thru Odds
    
        Set CurrentCkt = Range("CKT_" & CurrentCktNo & "_Poles")
        
        NoPoles = CurrentCkt.Value
        
        If IsNumeric(NoPoles) Then
            If NoPoles > SchdPoles Then NoPoles = 1
        Else
            NoPoles = 1
        End If
        
        If CurrentCktNo < 3 Then NoPoles = 1
        
        If CurrentCktNo > 2 And CurrentCktNo < 5 Then
            If NoPoles > 2 Then NoPoles = 1
        End If
        
        Select Case NoPoles
            Case 2, 3
            
            With Range(CurrentCkt.Address & ":" _
                       & CurrentCkt.Offset(-((2 * NoPoles) - 1), -1).Address)
                .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                
                If CurrentCktNo = 41 Then
                    .Borders(xlEdgeBottom).Weight = xlMedium
                Else
                    .Borders(xlEdgeBottom).Weight = xlThin
                End If
                
                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlInsideVertical).Weight = xlThin
            End With
            
            Case Else
            
            With Range(CurrentCkt.Address & ":" _
                       & CurrentCkt.Offset(-1, -1).Address)
                .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous

                If CurrentCktNo = 41 Then
                    .Borders(xlEdgeBottom).Weight = xlMedium
                Else
                    .Borders(xlEdgeBottom).Weight = xlThin
                End If

                .Borders(xlEdgeLeft).Weight = xlMedium
                .Borders(xlInsideVertical).Weight = xlThin
            End With
            
        End Select
    
    Next
    
    For CurrentCktNo = 2 To MaxCircuits Step 2 'Cycle Thru Evens
        
        Set CurrentCkt = Range("CKT_" & CurrentCktNo & "_Poles")
        
        NoPoles = CurrentCkt.Value
        
        If IsNumeric(NoPoles) Then
            If NoPoles > SchdPoles Then NoPoles = 1
        Else
            NoPoles = 1
        End If
        
        If CurrentCktNo < 3 Then NoPoles = 1
        
        If CurrentCktNo > 2 And CurrentCktNo < 5 Then
            If NoPoles > 2 Then NoPoles = 1
        End If
        
        Select Case CurrentCkt.Value
            Case 2, 3
            
            With Range(CurrentCkt.Address & ":" _
                       & CurrentCkt.Offset(-((2 * NoPoles) - 1), 1).Address)
                .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous

                If CurrentCktNo = 42 Then
                    .Borders(xlEdgeBottom).Weight = xlMedium
                Else
                    .Borders(xlEdgeBottom).Weight = xlThin
                End If

                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlEdgeRight).Weight = xlMedium
            End With
            
            Case Else
            
            With Range(CurrentCkt.Address & ":" _
                       & CurrentCkt.Offset(-1, 1).Address)
                .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous

                If CurrentCktNo = 42 Then
                    .Borders(xlEdgeBottom).Weight = xlMedium
                Else
                    .Borders(xlEdgeBottom).Weight = xlThin
                End If

                .Borders(xlInsideVertical).Weight = xlThin
                .Borders(xlEdgeRight).Weight = xlMedium
            End With

        End Select
    
    Next

    Call AutoHide
    
    ScreenUpdates (True)
    Application.EnableEvents = True
    
End Sub

Public Sub ToggleColor()
    
    Dim AffectRange As Range
    Dim Cell As Range
    
    MsgBox "TOGGLECOLOR EXECUTED"
    
    ScreenUpdates (False)
    InSub ("ON")
    
    Select Case GetInfo("SCHD_Type")
        Case "PANEL"
            Set AffectRange = Range("A1:M69")
        Case "BUS"
            Set AffectRange = Range("A1:G48")
    End Select
    
    If Range("A1").Interior.ColorIndex = 15 Then
    
        AffectRange.Interior.ColorIndex = xlColorIndexNone
        
    Else
                
        For Each Cell In AffectRange
        
            If Cell.Locked = True Then
                Cell.Interior.ColorIndex = 15
            Else
                Cell.Interior.ColorIndex = xlColorIndexNone
            End If
            
        Next
        
    End If

    InSub ("OFF")
    ScreenUpdates (True)
    
    ActiveSheet.Calculate

End Sub

