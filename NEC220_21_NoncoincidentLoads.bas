Attribute VB_Name = "NEC220_21_NoncoincidentLoads"
Option Explicit

Public Sub NoncoincidentLoadsDialog()

    NonCoincidentLoadsForm.Show

End Sub


Public Sub RemoveNoncoincidentLoads()

    Dim SchdPoles As Byte

    ScreenUpdates (False)
    Application.DisplayAlerts = False

    If HasNCLoads() Then
        ActiveWorkbook.Worksheets("Noncoincident Loads").Delete
        ActiveWorkbook.Worksheets(1).Range("A1").Activate
    End If

    SchdPoles = GetPoles()

'   ------- FIX TO ALLOW BUS OR PNL -------
'
'    With Range(Cells(56, 3).Address(False, False) & ":" _
'               & Cells(56, 8 + SchdPoles).Address(False, False))
'        .ClearContents
'        .Interior.ColorIndex = 15
'    End With

    Application.DisplayAlerts = True
    ScreenUpdates (True)

End Sub

Private Function HasNCLoads() As Boolean

    If IsMember(ActiveWorkbook.Worksheets, "Noncoincident Loads") Then
        HasNCLoads = True
    Else
        HasNCLoads = False
    End If

End Function

Private Sub SetupNCFormats(ByVal UpLeftCell As Excel.Range, GroupNo As Byte, NoLoads As Byte, NoSimul As Byte)

    Dim SchdPoles As Byte
    Dim CountPoles As Byte
    Dim PhaseNameRow As Byte
    Dim FormulaBuild As String
    Dim CountNoSimul As Byte
    
    PhaseNameRow = 11

    SchdPoles = GetPoles()
    
    UpLeftCell.Value = "Load Association Group " & GroupNo & _
                       "  [Where not more than (" & NoSimul & ") of the following (" & _
                       NoLoads & ") loads is likely to operate simultaneously.]"
                       
    UpLeftCell.Offset(1, 0).Value = "Load Description"
    UpLeftCell.Offset(1, 1).Value = "Ckt / Load No"
    UpLeftCell.Offset(1, 2).Value = "Schd Type"
    UpLeftCell.Offset(1, 3).Value = "Load Poles"
    UpLeftCell.Offset(1, 4).Value = "Path/Filename for Load Schedule"
    
    For CountPoles = 1 To SchdPoles
    
        FormulaBuild = "=-(sum(" & _
            Range(UpLeftCell.Offset(2, 4 + CountPoles), _
            UpLeftCell.Offset(1 + NoLoads, 4 + CountPoles)).Address & ")"
            
        For CountNoSimul = 1 To NoSimul
        
            FormulaBuild = FormulaBuild & "-large(" & _
                Range(UpLeftCell.Offset(2, 4 + CountPoles), _
                UpLeftCell.Offset(1 + NoLoads, 4 + CountPoles)).Address & _
                "," & CountNoSimul & ")"

        Next
                                                    
        FormulaBuild = FormulaBuild & ")"
                                                    
        UpLeftCell.Offset(0, 4 + CountPoles).Formula = FormulaBuild
    
        UpLeftCell.Offset(1, 4 + CountPoles).Formula = "=" & _
            Worksheets(GetSchdSht()).Cells(PhaseNameRow, 5 + CountPoles).Address(External:=True)
            
    Next

    With Range(UpLeftCell, UpLeftCell.Offset(0, 4 + SchdPoles))
        .Interior.ColorIndex = 35
        .Font.Bold = True
        .Font.Italic = True
    End With

    With Range(UpLeftCell.Offset(0, 5), UpLeftCell.Offset(0, 4 + SchdPoles))
        .NumberFormat = "0_);[Red](0)"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        If SchdPoles > 1 Then .Borders(xlInsideVertical).LineStyle = xlContinuous
    
    End With

    With Range(UpLeftCell.Offset(1, 0), UpLeftCell.Offset(NoLoads + 1, 4 + SchdPoles))
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With

End Sub

Private Function FindNCStartCell(Optional GroupNo As Byte = 0) As Excel.Range

    Dim Found As Boolean
    Dim TestCell As Excel.Range
    Dim CurrentRow As Double
    Dim NoLoads As Byte
    
    CurrentRow = 4

    Do
    
        Set TestCell = ActiveWorkbook.Worksheets("NonCoincident Loads").Cells(CurrentRow, 2)
        
        If TestCell.Text <> "" Then
        
            If Mid(TestCell.Value, 24, 1) = GroupNo Then Found = True
        
            NoLoads = Mid(TestCell.Value, InStr(1, TestCell.Value, ") loads") - 1, 1)

            CurrentRow = CurrentRow + 3 + NoLoads
        
        Else
            Found = True
        End If
    
    Loop Until Found = True
    
    Set FindNCStartCell = TestCell

End Function

Private Function FindNextGroupNo() As Byte

    Dim Found As Boolean
    Dim TestCell As Excel.Range
    Dim CurrentRow As Double
    Dim NoLoads As Byte
    Dim GroupNo As Byte
    
    GroupNo = 0
    CurrentRow = 4

    Do
        GroupNo = GroupNo + 1
    
        Set TestCell = ActiveWorkbook.Worksheets("NonCoincident Loads").Cells(CurrentRow, 2)
        
        If TestCell.Text <> "" Then
        
            NoLoads = Mid(TestCell.Value, InStr(1, TestCell.Value, ") loads") - 1, 1)
            
            CurrentRow = CurrentRow + 3 + NoLoads
        
        Else
            Found = True
        End If
    
    Loop Until Found = True
    
    FindNextGroupNo = GroupNo

End Function

Private Sub FixGroupNumbers()

    Dim Found As Boolean
    Dim TestCell As Excel.Range
    Dim CurrentRow As Double
    Dim NoLoads As Byte
    Dim GroupNo As Byte
    
    GroupNo = 0
    CurrentRow = 4

    Do
        GroupNo = GroupNo + 1
    
        Set TestCell = ActiveWorkbook.Worksheets("NonCoincident Loads").Cells(CurrentRow, 2)
        
        If TestCell.Text <> "" Then
        
            If Mid(TestCell.Value, 24, 1) <> GroupNo Then
                TestCell.Value = Replace(TestCell.Value, _
                                    Find:="Group " & Mid(TestCell.Value, 24, 1), _
                                    Replace:="Group " & GroupNo)
            End If
        
            NoLoads = Mid(TestCell.Value, InStr(1, TestCell.Value, ") loads") - 1, 1)
            
            CurrentRow = CurrentRow + 3 + NoLoads
        
        Else
            Found = True
        End If
    
    Loop Until Found = True

End Sub

Private Sub DeleteGroup(GroupNo As Byte)
    
    Dim UpLeftCell As Excel.Range
    Dim NoLoads As Byte
    
    UpLeftCell = FindNCStartCell(GroupNo)
    
    NoLoads = Mid(UpLeftCell.Value, InStr(1, UpLeftCell.Value, ") loads") - 1, 1)
    
    Range(UpLeftCell, UpLeftCell.Offset(2 + NoLoads, 7)).Delete (xlShiftUp)

End Sub

Private Sub ReviseGroup(GroupNo As Byte, NewNoLoads As Byte, NewNoSimul As Byte)

    Dim UpLeftCell As Excel.Range
    Dim OldNoLoads As Byte
    Dim CountLines As Byte
    Dim WhichLoad As Byte
    
    Set UpLeftCell = FindNCStartCell(GroupNo)
    
    OldNoLoads = Mid(UpLeftCell.Value, InStr(1, UpLeftCell.Value, ") loads") - 1, 1)
    
    If NewNoLoads > OldNoLoads Then
    
        For CountLines = 1 To (NewNoLoads - OldNoLoads)
    
            Range(UpLeftCell.Offset(2 + OldNoLoads, 0), _
                  UpLeftCell.Offset(2 + OldNoLoads, 7)).Insert (xlShiftDown)
        
        Next
    
    Else
    
        For CountLines = 1 To (OldNoLoads - NewNoLoads)
        
            WhichLoad = Application.InputBox("Remove Which Load?")
    
            Range(UpLeftCell.Offset(1 + WhichLoad, 0), _
                  UpLeftCell.Offset(1 + WhichLoad, 7)).Delete (xlShiftUp)
        
        Next
    
    End If
    
    Call SetupNCFormats(UpLeftCell, GroupNo, NewNoLoads, NewNoSimul)

End Sub

Private Sub BuildNCSummation()

    Dim FirstNCLoadCell As Excel.Range
    Dim CountPoles As Byte
    Dim Found As Boolean
    Dim TestCell As Excel.Range
    Dim CurrentRow As Double
    Dim NoLoads As Byte
    Dim FormulaBuild As String
    
    Select Case GetInfo("SCHD_Type")
    Case "PANEL"
        Set FirstNCLoadCell = Worksheets(GetSchdSht()).Range("Misc2_L1_VA")
    Case "BUS"
        Set FirstNCLoadCell = Worksheets(GetSchdSht()).Range("Misc1_L1_VA")
    End Select
    
    Worksheets(GetSchdSht()).Cells(FirstNCLoadCell.Row, 3).Value = "Reduction for NEC 220.21 NonCoincident Loads"
    
    For CountPoles = 1 To GetPoles()
      
        FormulaBuild = "="
        
        CurrentRow = 4
        
        Found = False
        
        Do
        
            Set TestCell = ActiveWorkbook.Worksheets("NonCoincident Loads").Cells(CurrentRow, 2)
            
            If TestCell.Text <> "" Then
           
                If FormulaBuild <> "=" Then FormulaBuild = FormulaBuild & "+"
                
                FormulaBuild = FormulaBuild & TestCell.Offset(0, 4 + CountPoles).Address(External:=True)
            
                NoLoads = Mid(TestCell.Value, InStr(1, TestCell.Value, ") loads") - 1, 1)
    
                CurrentRow = CurrentRow + 3 + NoLoads
            
            Else
                
                Found = True
            
            End If
        
        Loop Until Found = True
        
        With FirstNCLoadCell.Offset(0, CountPoles - 1)
            .Formula = FormulaBuild
            .Interior.ColorIndex = 8
        End With
     
    Next

End Sub


Private Sub test3()

    Call ReviseGroup(Application.InputBox("Group"), _
        Application.InputBox("New No Loads"), _
        Application.InputBox("New No Simul"))

End Sub

Private Sub test()

    Call SetupNCFormats(FindNCStartCell, FindNextGroupNo, _
        Application.InputBox("No Loads"), _
        Application.InputBox("No Simul"))

End Sub

Private Sub test2()

    Call DeleteGroup(Application.InputBox("group to delete"))
    Call FixGroupNumbers

End Sub
