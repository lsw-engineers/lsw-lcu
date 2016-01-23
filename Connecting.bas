Attribute VB_Name = "Connecting"
Option Explicit

Public Sub AddConnectionDialog()

    AddConnectionForm.Show

End Sub

Public Sub LinkSchedule(FileToLink As String, ConnectionType As String, _
                        NoPoles As Byte, FirstCktNo As Byte)
    
    'Acceptable ConnectionType is "CKT" "Misc1" "Misc2" "Load1" "Load2" etc.
    'FirstCktNo will be an integer 1 to 84
    'NoPoles will be an integer 1 to 3
    
    ScreenUpdates (False)

    Dim AssocCkts As Variant

    Select Case NoPoles

        Case 1 And ConnectionType = "CKT"
            AssocCkts = Array(FirstCktNo)
        Case 2 And ConnectionType = "CKT"
            AssocCkts = Array(FirstCktNo, FirstCktNo + 2)
        Case 3 And ConnectionType = "CKT"
            AssocCkts = Array(FirstCktNo, FirstCktNo + 2, FirstCktNo + 4)
    
        Case 1 And ConnectionType <> "CKT" And FirstCktNo = 1
            AssocCkts = Array("L1")
        Case 2 And ConnectionType <> "CKT" And FirstCktNo = 1
            AssocCkts = Array("L1", "L2")
        Case 3 And ConnectionType <> "CKT" And FirstCktNo = 1
            AssocCkts = Array("L1", "L2", "L3")

        Case 1 And ConnectionType <> "CKT" And FirstCktNo = 2
            AssocCkts = Array("L2")
        Case 2 And ConnectionType <> "CKT" And FirstCktNo = 2
            AssocCkts = Array("L2", "L3")
        Case 3 And ConnectionType <> "CKT" And FirstCktNo = 2
            AssocCkts = Array("L2", "L3", "L1")

        Case 1 And ConnectionType <> "CKT" And FirstCktNo = 3
            AssocCkts = Array("L3")
        Case 2 And ConnectionType <> "CKT" And FirstCktNo = 3
            AssocCkts = Array("L3", "L1")
        Case 3 And ConnectionType <> "CKT" And FirstCktNo = 3
            AssocCkts = Array("L3", "L1", "L2")
        
        Case Else
        MsgBox "No Cases Matched"

    End Select
    
    Dim CurrentLine As Byte
    CurrentLine = 1
    
    Dim CellTypes As Variant
    Dim CurrentCkt As Variant
    Dim CurrentType As Variant

    CellTypes = Array("VA", "C", "L", "M", "ML", "R", "H", "T", "K", "KQ", _
                            "X", "XL", "XS", "XQ", "Z", "ZL", "ZS", "ZQ")
        
    For Each CurrentCkt In AssocCkts

        For Each CurrentType In CellTypes

            Dim CellToModify As String
            CellToModify = ConnectionType & "_" & CurrentCkt & "_" & CurrentType

            Dim UseFormula As String
            
            UseFormula = "='" & FileToLink & "'!Total_L" & CurrentLine & "_" & CurrentType
            
            Range(CellToModify).Formula = UseFormula

        Next
        
        CurrentLine = CurrentLine + 1

    Next
    
    If ConnectionType <> "CKT" And GetInfo("SCHD_Type") = "PANEL" Then
    
        Dim SchdPoles As Byte
        SchdPoles = GetPoles()
        
        Dim Row As Long
        Row = Range(ConnectionType & "_L1_VA").Row
        
        Cells(Row, 6 + SchdPoles).Value = "VIA FEED-THRU LUGS"
    
    End If
    
    ScreenUpdates (True)
    
End Sub

