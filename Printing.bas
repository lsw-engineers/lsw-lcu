Attribute VB_Name = "Printing"
Option Explicit

Public Sub PrintAllSchds()

    Dim LCUWkSht As Excel.Worksheet
    Dim StartWkbk As Excel.Workbook
    Dim CurrWkbk As Excel.Workbook
    Dim LastLine As Integer
    Dim CurrRow As Integer
    Dim CurrFileName As String
    Dim StartFileName As String
    Dim PrtChoice As Byte
    
    Set LCUWkSht = Workbooks("lcu.xla").Sheets(1)
    
    Set StartWkbk = ActiveWorkbook
    
    StartFileName = StartWkbk.Path & "\" & StartWkbk.Name
       
    LastLine = ListFiles(StartWkbk)

    PrtChoice = MsgBox("You are about to print " & LastLine - 2 _
                        & " schedules.  Are you sure?", _
                        vbExclamation + vbYesNo, "Print Confirm")

    If PrtChoice = vbYes Then
    
        For CurrRow = 3 To LastLine
            CurrFileName = LCUWkSht.Cells(CurrRow, 2).Text & _
                           LCUWkSht.Cells(CurrRow, 3).Text
    
            If CurrFileName <> StartFileName Then
                Set CurrWkbk = Workbooks.Open(CurrFileName, 0)
                
                If IsMember(CurrWkbk.Names, "LCU_Version") Then
                    CurrWkbk.Sheets(1).PrintOut
                End If
                
                CurrWkbk.Close (True)
                
            Else
            
                Set CurrWkbk = StartWkbk
                
                If IsMember(StartWkbk.Names, "LCU_Version") Then
                    CurrWkbk.Sheets(1).PrintOut
                End If
            
            End If
        Next
        
    End If
    
    LCUWkSht.Range("B3", LCUWkSht.Range("C3").End(xlDown)).ClearContents

End Sub

Private Function ListFiles(ThisWkbk As Excel.Workbook) As Integer

    Dim LCUWkSht As Excel.Worksheet
    Dim Folder As String
    Dim CurrRow As Integer
    Dim Count As Integer
    
    Set LCUWkSht = Workbooks("lcu.xla").Sheets(1)
    
    CurrRow = 3
    
    Folder = ThisWkbk.Path
    
    With Application.FileSearch
        .NewSearch
        .LookIn = Folder
        .FileName = "*.xls"
        .SearchSubFolders = True
        .Execute
        For Count = 1 To .FoundFiles.Count
        
            If GetInfo("LCU_Version", .FoundFiles(Count)) <> "INVALID" Then
            
                LCUWkSht.Cells(CurrRow, 2) = PathOnly(.FoundFiles(Count))
                LCUWkSht.Cells(CurrRow, 3) = FileNameOnly(.FoundFiles(Count))
                CurrRow = CurrRow + 1
            
            End If
            
        Next Count
    End With
    
    LCUWkSht.Range("B3:C" & CurrRow - 1).Sort _
        Key1:=LCUWkSht.Cells(3, 2), _
        Key2:=LCUWkSht.Cells(3, 3)
    
    ListFiles = CurrRow - 1
    
End Function

