Attribute VB_Name = "formulasSubroutines"
Sub ConvertToValues(control As IRibbonControl)
Dim myRange As Range
Dim myCell As Range
Select Case MsgBox("You Can't Undo This Action. " & "Save Workbook First?", vbYesNoCancel, "Alert")
Case Is = vbYes
ThisWorkbook.Save
Case Is = vbCancel
Exit Sub
End Select
Set myRange = Selection
For Each myCell In myRange
If myCell.HasFormula Then
myCell.Formula = myCell.Value
End If
Next myCell
End Sub
'It will check your selection and then remove all the extra spaces from that.
Sub RemoveSpaces(control As IRibbonControl)
Dim myRange As Range
Dim myCell As Range
Select Case MsgBox("You Can't Undo This Action. " & "Save Workbook First?", vbYesNoCancel, "Alert")
Case Is = vbYesThisWorkbook.Save
Case Is = vbCancel
Exit Sub
End Select
Set myRange = Selection
For Each myCell In myRange
If Not IsEmpty(myCell) Then
myCell = Trim(myCell)
End If
Next myCell
End Sub

Sub degreeSymbol(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection
rng.Select
If ActiveCell <> "" Then
If IsNumeric(ActiveCell.Value) Then
ActiveCell.Value = ActiveCell.Value & ChrW(176)
End If
End If
Next
End Sub

Sub getSquareRoot(control As IRibbonControl)
Dim rng As Range
Dim i As Integer
For Each rng In Selection
If WorksheetFunction.IsNumber(rng) Then
rng.Value = Sqr(rng)
Else
End If
Next rng
End Sub
Sub getCubeRoot(control As IRibbonControl)
Dim rng As Range
Dimi As Integer
For Each rng In Selection
If WorksheetFunction.IsNumber(rng) Then
rng.Value = rng ^ (1 / 3)
Else
End If
Nextrng
End Sub
Sub addcAlphabets(control As IRibbonControl)
Dim i As Integer
For i = 65 To 90
ActiveCell.Value = Chr(i)
ActiveCell.Offset(1, 0).Select
Next i
End Sub
Sub addsAlphabets(control As IRibbonControl)
Dim i As Integer
For i = 97 To 122
ActiveCell.Value = Chr(i)
ActiveCell.Offset(1, 0).Select
Next i
End Sub
Sub convertToNumbers(control As IRibbonControl)
Dim rng As Range
Selection.Value = Selection.Value
For Each rng In Selection
If Not WorksheetFunction.IsNonText(rng) Then
rng.Value = WorksheetFunction.Arabic(rng)
End If
Next rng
End Sub

Sub removeNegativeSign(control As IRibbonControl)
Dim rng As Range
Selection.Value = Selection.Value
For Each rng In Selection
If Worksheet.Function.IsNumber(rng) Then rng.Value = Abs(rng)
End If
Next rng
End Sub
Sub replaceBlankWithZero(control As IRibbonControl)
Dim rng As Range
Selection.Value = Selection.Value
For Each rng In Selection
If rng = "" Or rng = " " Then
rng.Value = "0"
Else
End If
Next rng
End Sub
