Attribute VB_Name = "Admin"
Option Explicit

Sub About_LCU()
    LCU_About.Show
End Sub

Sub Setup_OnKey()
    Application.OnKey "%{F5}", "SetCktDivisions"
End Sub

Sub Select_Print_Area()
    Dim SchdPoles As Byte
    SchdPoles = GetPoles()
    
    Range(Cells(8, 2).Address(False, False) & ":" _
           & Cells(67, 9 + SchdPoles).Address(False, False)).Name = "Print_Area"
    Range("Print_Area").Select
End Sub

Public Sub ExportAllNames()

    ScreenUpdates (False)

    Dim CurrentRow As Double
   
    Dim ExportFrom As Excel.Workbook
    Set ExportFrom = ActiveWorkbook
    
    Dim ExportSheet As Excel.Worksheet
    Set ExportSheet = Workbooks("Cell Names.xls").Worksheets("Export")
    
    ExportSheet.Range("A1").EntireColumn.NumberFormat = "@"
    ExportSheet.Range("B1").EntireColumn.NumberFormat = "@"
    
    Dim CurrentName As Excel.Name
    
    CurrentRow = 0
    
    For Each CurrentName In ExportFrom.Names
     
        CurrentRow = CurrentRow + 1
        
        ExportSheet.Range("A" & CurrentRow).Value = "'" & CurrentName.Name
        ExportSheet.Range("B" & CurrentRow).Value = "'" & CurrentName.RefersToLocal

    Next

    ScreenUpdates (True)
    ExportSheet.Parent.Activate
    
End Sub

Public Sub CleanUpNames(Optional Quiet As Boolean = False)

    '*** DELETES ALL NAMED RANGES IN THE WORKBOOK.  ANY REFERENCES TO THE   ***
    '*** DELETED NAMES WILL RESULT IN A #NAME? ERROR.                       ***
    
    Dim n As Name
    Dim i As Double
    
    i = 0
    
    For Each n In ActiveWorkbook.Names
        If InStr(1, n.RefersTo, "#REF!") <> 0 Then
        
            n.Delete
            i = i + 1
        
        End If
        
    Next n
    
    If Quiet = False Then MsgBox i & " unreferenced names were removed."

End Sub

Public Sub DeleteAllNames(Optional Quiet As Boolean = False)

    '*** DELETES ALL NAMED RANGES IN THE WORKBOOK.  ANY REFERENCES TO THE   ***
    '*** DELETED NAMES WILL RESULT IN A #NAME? ERROR.                       ***
    
    Dim n As Name
    Dim Count As Integer
    For Each n In ActiveWorkbook.Names
        n.Delete
        Count = Count + 1
    Next n
    If Quiet = False Then MsgBox "All names were deleted!"
End Sub

Public Sub DefineAllNames()

    '*** CREATES NAMED RANGES BASED ON A MASTER TABLE IN "CELL NAMES.XLS    ***
    Dim i As Integer
    Dim Schd_Type As String
    
    Schd_Type = Application.InputBox("PANEL OR BUS?")
    
    ScreenUpdates (False)
    
    Select Case Schd_Type
    
    Case "PANEL"
    
        For i = 1 To Workbooks("Cell Names.xls").Worksheets("Master List").Range("K1").Value
            ActiveWorkbook.Names.Add _
            Name:=Workbooks("Cell Names.xls").Worksheets("Master List").Range("D" & i).Value, _
            RefersTo:=Workbooks("Cell Names.xls").Worksheets("Master List").Range("F" & i).Value
            'Range(Workbooks("Cell Names.xls").Worksheets("Master List").Range("G" & i).Value). _
                Interior.ColorIndex = xlColorIndexNone
        
        Next
        
    Case "BUS"
    
        For i = 1 To Workbooks("Cell Names.xls").Worksheets("Master Bus").Range("M1").Value
            ActiveWorkbook.Names.Add _
            Name:=Workbooks("Cell Names.xls").Worksheets("Master Bus").Range("D" & i).Value, _
            RefersTo:=Workbooks("Cell Names.xls").Worksheets("Master Bus").Range("H" & i).Value
            'Range(Workbooks("Cell Names.xls").Worksheets("Master Bus").Range("I" & i).Value). _
                Interior.ColorIndex = 4
        
        Next
    
    End Select
    
    ScreenUpdates (True)
    
End Sub

Public Sub ScreenUpdates(Setting As Boolean)
    
    Static LastSetting As Boolean
    
    Select Case Setting

    Case True 'Turning it On
    
        If LastSetting And Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    
    Case False 'Turning it Off
    
        LastSetting = Application.ScreenUpdating
        
        If Application.ScreenUpdating Then Application.ScreenUpdating = False
    
    End Select

End Sub

Public Function FileExists(fname) As Boolean
'   Returns TRUE if the file exists
    Dim X As String
    X = Dir(fname)
    If X <> "" Then FileExists = True _
        Else FileExists = False
End Function

Public Function FileNameOnly(ByVal pname As String) As String
'   Returns the filename from a path/filename string
    Dim i As Integer, length As Integer, temp As String
    length = Len(pname)
    temp = ""
    For i = length To 1 Step -1
        If Mid(pname, i, 1) = Application.PathSeparator Then
            FileNameOnly = temp
            Exit Function
        End If
        temp = Mid(pname, i, 1) & temp
    Next i
    FileNameOnly = pname
End Function

Public Function PathOnly(ByVal PathFileName As String) As String
    
    Dim PathOnlyLength As Long
    Dim FileName As String
    
    FileName = FileNameOnly(PathFileName)
    
    PathOnlyLength = Len(PathFileName) - Len(FileName)
    
    PathOnly = Left(PathFileName, PathOnlyLength)

End Function


Public Function InSub(Setting As String) As Boolean
    
    'Setting should be "ON" or "OFF" to set or "QUERY" for return state
    
    Static CurrentState As Boolean
    
    Select Case Setting

    Case "ON"

        CurrentState = True
    
    Case "OFF"
    
        CurrentState = False
    
    Case "QUERY"
    
        InSub = CurrentState

    End Select

End Function

Public Function IsMember(Collection As Object, Item As String) As Boolean

    Dim Obj As Object
    
    On Error Resume Next
    
    Set Obj = Collection(Item)
    
    IsMember = Not Obj Is Nothing

End Function

Public Function GetPoles(Optional Wb As Excel.Workbook = Nothing) As Byte

    If Wb Is Nothing Then
        GetPoles = Left(GetInfo("SCHD_Poles"), 1)
    Else
        GetPoles = Left(GetInfo("SCHD_Poles", Wb.Name), 1)
    End If

End Function

Public Function GetSchdSht() As Byte
    
    On Error Resume Next
    
    Dim Wksht As Excel.Worksheet
    Dim Trash As String
    
    Trash = ""
    GetSchdSht = 0

    For Each Wksht In ActiveWorkbook.Worksheets

        Trash = Wksht.Range("LCU_Version").Value
        
        If Trash <> "" Then
        
            GetSchdSht = Wksht.Index
            
            Exit Function
            
        End If
        
    Next

End Function

Public Function GetInfo(InfoType As String, Optional FilePathName As String = "") As String

    Dim Ref_Cell As Excel.Range
    Dim FileName As String
    
    FileName = FileNameOnly(FilePathName)

    If FilePathName = "" Then ' If none specified, assume the ActiveWorkbook
    
        If IsMember(ActiveWorkbook.Names, InfoType) Then
        
            GetInfo = ActiveWorkbook.Sheets(1).Range(InfoType).Value
        
        Else
            
'            MsgBox "Invalid or Incompatable Load Schedule:" & vbCrLf & vbCrLf & _
                    ActiveWorkbook.Name , vbExclamation
                    
            GetInfo = "INVALID"
        
        End If
    
    Else    ' If a filename was specified ...
    
        If IsMember(Workbooks, FileName) Then
            
            If IsMember(Workbooks(FileName).Names, InfoType) Then
            
                GetInfo = Workbooks(FileName).Sheets(1).Range(InfoType).Value
            
            Else
                
'                MsgBox "Invalid or Incompatable Load Schedule:" & vbCrLf & vbCrLf & _
                        FileName, vbExclamation
                        
                GetInfo = "INVALID"
            
            End If
            
        Else  ' If the specified file is closed...
            
            On Error GoTo ErrorClosed
            
            Set Ref_Cell = Workbooks("lcu.xla").Worksheets(1).Range("A1")
            Ref_Cell.Formula = "='" & FilePathName & "'!" & InfoType
            
            GetInfo = Ref_Cell.Value
            
            Ref_Cell.ClearContents
        
        End If
        
    End If
    
    Exit Function
    
ErrorClosed:

    Ref_Cell.ClearContents
'    MsgBox "Invalid or Incompatable Load Schedule:" & vbCrLf & vbCrLf & _
            FilePathName, vbExclamation
            
    GetInfo = "INVALID"
    
End Function

Public Sub RunSpanner()

    Application.EnableEvents = False
    
    Application.Run "span12.xla!Spanner_Go"

    Application.EnableEvents = True
    
End Sub

Public Sub DeleteSpannerNames()

    If IsMember(ActiveWorkbook.Names, "Auto_Close_Spanner") Then _
        ActiveWorkbook.Names("Auto_Close_Spanner").Delete
    
    If IsMember(ActiveWorkbook.Names, "Spanner_Auto_File") Then _
        ActiveWorkbook.Names("Spanner_Auto_File").Delete
        
    If IsMember(ActiveWorkbook.Names, "Spanner_Auto_Select") Then _
        ActiveWorkbook.Names("Spanner_Auto_Select").Delete
    
End Sub
