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


Sub AutoFitColumns(control As IRibbonControl)
Cells.Select
Cells.EntireColumn.AutoFit
End Sub

Sub AutoFitRows(control As IRibbonControl)
Cells.Select
Cells.EntireRow.AutoFit
End Sub

Sub AutoFitRowsColumns(control As IRibbonControl)
AutoFitColumns
AutoFitRows
End Sub


Sub TimeStamp(control As IRibbonControl)
Dim i As Integer
For i = 1 To 24
ActiveCell.FormulaR1C1 = i & ":00"
ActiveCell.NumberFormat = "[$-409]h:mm AM/PM;@"
ActiveCell.Offset(RowOffset:=1, ColumnOffset:=0).Select
Next i
End Sub

Sub date2day(control As IRibbonControl)
Dim tempCell As Range
Selection.Value = Selection.Value
For Each tempCell In Selection
If IsDate(tempCell) = True Then
With tempCell
.Value = Day(tempCell)
.NumberFormat = "0"
End With
End If
Next tempCell
End Sub
Sub date2year(control As IRibbonControl)
Dim tempCell As Range
Selection.Value = Selection.Value
For Each tempCell In Selection
If IsDate(tempCell) = True Then
With tempCell
.Value = Year(tempCell)
.NumberFormat = "0"
End With
End If
Next tempCell
End Sub
Sub removeTime(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection
If IsDate(rng) = True Then
rng.Value = VBA.Int(rng.Value)
End If
Next
Selection.NumberFormat = "dd-mmm-yy"
End Sub
Sub removeDate(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection
If IsDate(rng) = True Then
rng.Value = rng.Value - VBA.Fix(rng.Value)
End If
NextSelection.NumberFormat = "hh:mm:ss am/pm"
End Sub

