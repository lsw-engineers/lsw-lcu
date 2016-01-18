Attribute VB_Name = "LocalFunctions"
Option Explicit

Public Function MAXDEV(GivenRange As Excel.Range) As Double

    Application.Volatile

    MAXDEV = Application.Run("lcu.xla!lcumaxdev", GivenRange)
    
End Function

Public Function MAXDEVPHASE(GivenRange As Excel.Range) As String
    
    Application.Volatile

    MAXDEVPHASE = Application.Run("lcu.xla!lcumaxdevphase", GivenRange)
    
End Function



