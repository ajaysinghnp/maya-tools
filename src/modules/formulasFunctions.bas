Attribute VB_Name = "formulasFunctions"
Public Function rvrse(ByVal cell As Range) As String
rvrse = VBA.StrReverse(cell.Value)
End Function
