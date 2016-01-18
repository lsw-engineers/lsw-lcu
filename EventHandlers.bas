Attribute VB_Name = "EventHandlers"
Option Explicit

Dim X As New AppEvents

Public Sub InitAppEvents()

    Call Setup_OnKey

    Set X.XL = Application

End Sub

Public Sub Schd_Calculate(Sht As Excel.Worksheet)

    If UCase(Sht.Parent.Name) = UCase("lcu.xla") Then Exit Sub
    
'    If Sht.ProtectContents = True Then _   ****MOVED TO APP EVENTS***
'            Sht.Protect UserInterfaceOnly:=True

    Application.EnableEvents = False

    Call IsOverLoaded(Sht)

    Application.EnableEvents = True

End Sub

Sub Schd_Change(ByVal Target As Excel.Range)

    Dim ValidRange As Range
    Dim SchdPoles As Byte

    On Error Resume Next
    
    SchdPoles = GetPoles()
    
    Select Case GetInfo("SCHD_Type")
        
        Case "PANEL"
            
            Set ValidRange = Range("B12", Cells(56, 9 + SchdPoles))
            
        Case "BUS"
        
            Set ValidRange = Range("B10", Cells(35, 3 + SchdPoles))
        
    End Select

    Application.EnableEvents = False
    
    If Not InSub("QUERY") Then ScreenUpdates (False)
                
    If Not Intersect(Target, ValidRange) Is Nothing Then
    
        Call AutoHide
            
    End If
    
    If Not InSub("QUERY") Then ScreenUpdates (True)
    
    Application.EnableEvents = True
              
End Sub


Public Sub AutoHide()

    Dim Row As Integer
    Dim Column As Integer
    Dim RowInUse As Boolean
    Dim FirstCol As Integer
    Dim LastCol As Integer
    Dim FirstRow As Integer
    Dim LastRow As Integer
    Dim SchdPoles As Byte
    
    On Error Resume Next

    SchdPoles = GetPoles()
    
    Select Case GetInfo("SCHD_Type")
        
        Case "PANEL"
  
            FirstCol = 3
            LastCol = 8 + SchdPoles
            
            FirstRow = 54
            LastRow = 65
            
        Case "BUS"
           
            FirstCol = 2
            LastCol = 3 + SchdPoles
            
            FirstRow = (Range("Misc1_LT").Row + 2)
            LastRow = 43
        
    End Select

    For Row = FirstRow To LastRow 'Auto-Hide Demand Factors, etc.
    
        RowInUse = False
        
        For Column = FirstCol To LastCol
                   
            If Cells(Row, Column).Value <> 0 Then
            
                RowInUse = True
                Exit For
            
            End If
            
        Next
            
        If RowInUse And Range("A" & Row).EntireRow.Hidden = True Then
            
            Range("A" & Row).EntireRow.Hidden = False
                    
        ElseIf Not RowInUse And Not Range("A" & Row).EntireRow.Hidden = True Then
            
            Range("A" & Row).EntireRow.Hidden = True
        
        End If
        
    Next

End Sub
