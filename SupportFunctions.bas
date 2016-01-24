Attribute VB_Name = "SupportFunctions"
Public Function isEven(Number As Byte) As Byte

    isEven = 1 - (Number Mod 2)

End Function

Public Function XLMod(a, b)
    ' This replicates the Excel MOD function
    XLMod = a - b * Int(a / b)
End Function
