Attribute VB_Name = "TextToNumber"
Sub ExtrNumbersFromRange(control As IRibbonControl)
    Dim xRg As Range
    Dim xDRg As Range
    Dim xRRg As Range
    Dim nCellLength As Integer
    Dim xNumber As Integer
    Dim strNumber As String
    Dim xTitleId As String
    Dim xI As Integer
    xTitleId = "MayaTools Extract Numbers"
    Set xDRg = Application.InputBox("Please select the Range:", xTitleId, "", Type:=8)
    If TypeName(xDRg) = "Nothing" Then Exit Sub
    Set xRRg = Application.InputBox("Please select output cell:", xTitleId, "", Type:=8)
    If TypeName(xRRg) = "Nothing" Then Exit Sub
    xI = 0
    strNumber = ""
  For Each xRg In xDRg
    xI = xI + 1
    nCellLength = Len(xRg)
    For xNumber = 1 To nCellLength
      If IsNumeric(Mid(xRg, xNumber, 1)) Or Mid(xRg, xNumber, 1) = "." Then
        strNumber = strNumber & Mid(xRg, xNumber, 1)
      End If
    Next xNumber
    xRRg.Item(xI) = strNumber
    strNumber = ""
  Next xRg
End Sub
