Attribute VB_Name = "HelperSubroutines"
Sub AddSerialNumbers(control As IRibbonControl)
'automatically add the serial numbers to the sheet
Dim i As Integer
On Error GoTo Last
i = InputBox("Enter Value", "MayaTools for Excel")
For i = 1 To i
ActiveCell.Value = i
ActiveCell.Offset(1, 0).Activate
Next i
Last: Exit Sub
End Sub

Sub InsertMultipleColumns(control As IRibbonControl)
'Insert multiple columns to the sheets at once
Dim i As Integer
Dim j As Integer
ActiveCell.EntireColumn.Select
On Error GoTo Last
i = InputBox("Enter number of columns to insert", "MayaTools for Excel")
For j = 1 To i
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightorAbove
Next j
Last: Exit Sub
End Sub

Sub InsertMultipleRows(control As IRibbonControl)
'Insert multiple rows to the sheets at once
Dim i As Integer
Dim j As Integer
ActiveCell.EntireRow.Select
On Error GoTo Last
i = InputBox("Enter number of columns to insert", "MayaTools for Excel")
For j = 1 To i
Selection.Insert Shift:=xlToDown, CopyOrigin:=xlFormatFromRightorAbove
Next j
Last: Exit Sub
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

