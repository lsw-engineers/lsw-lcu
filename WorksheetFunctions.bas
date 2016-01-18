Attribute VB_Name = "WorksheetFunctions"
Option Explicit

Function lcumaxdev(GivenRange As Excel.Range) As Double
    
    '*** WORKSHEET FUNCTION: RETURNS THE MAXIMUM    ***
    '*** DEVIATION FROM A GIVEN RANGE               ***
    
    Application.Volatile
    
    Dim Total As Double
    Dim Average As Double
    Dim MaxDeviation As Double
    Dim NoCells As Byte
        
    Total = 0
    Average = 0
    MaxDeviation = 0
    NoCells = 0
    
    Dim Cell As Excel.Range
    
    For Each Cell In GivenRange
    
        Total = Total + Cell.Value
        NoCells = NoCells + 1
    
    Next
    
    Average = Total / NoCells
    
    For Each Cell In GivenRange
    
        If Abs(Cell.Value - Average) > MaxDeviation Then
            MaxDeviation = Abs(Cell.Value - Average)
        End If
    
    Next
               
    lcumaxdev = MaxDeviation
    
End Function

Function lcumaxdevphase(GivenRange As Excel.Range) As String
    
    '*** WORKSHEET FUNCTION: RETURNS THE "PHASE" OF THE MAXIMUM DEVIATION   ***
    '*** BASED ON A FIXED ROW NUMBER FOR THE "PHASE" DESCRIPTION". RETURNS  ***
    '*** THE VALUE OF THE CELL.                                             ***
    
    Dim PhaseNameRow As Integer
    
    Select Case GetInfo("SCHD_Type", GivenRange.Parent.Parent.Name)
        
        Case "PANEL"
            
             PhaseNameRow = 11
        
        Case "BUS"
             
             PhaseNameRow = 8
    
    End Select
    
    Application.Volatile
    
    Dim Total As Double
    Dim Average As Double
    Dim MaxDeviation As Double
    Dim NoCells As Byte
    Dim MaxDevColumn As Integer
        
    Total = 0
    Average = 0
    MaxDeviation = 0
    NoCells = 0
    MaxDevColumn = 0
    
    Dim Cell As Excel.Range
    
    For Each Cell In GivenRange
    
        Total = Total + Cell.Value
        NoCells = NoCells + 1
    
    Next
    
    Average = Total / NoCells
    
    For Each Cell In GivenRange
    
        If Abs(Cell.Value - Average) > MaxDeviation Then
            MaxDeviation = Abs(Cell.Value - Average)
            MaxDevColumn = Cell.Column
        End If
    
    Next

    If MaxDevColumn = 0 Then
        lcumaxdevphase = "BALANCED"
    Else
        lcumaxdevphase = Cells(PhaseNameRow, MaxDevColumn).Value
    End If
    
End Function
