Attribute VB_Name = "pivotTablesSubroutines"
Sub HideSubtotals(control As IRibbonControl)
Dim pt As PivotTable
Dim pf As PivotField
On Error Resume Next
Set pt = ActiveSheet.PivotTables(ActiveCell.PivotTable.Name)
If pt Is Nothing Then
MsgBox "You must place your cursor inside of a PivotTable."
Exit Sub
End If
For Each pf In pt.PivotFields
pf.Subtotals(1) = True
pf.Subtotals(1) = False
Next pf
End Sub

Sub CloseAllWorkbooks(control As IRibbonControl)
Dim wbs As Workbook
For Each wbs In Workbooks
wbs.Close SaveChanges:=True
Next wb
End Sub
'Auto Update Pivot Table Range
Sub UpdatePivotTableRange(control As IRibbonControl)
Dim Data_Sheet As Worksheet
Dim Pivot_Sheet As Worksheet
Dim StartPoint As Range
Dim DataRange As Range
Dim PivotName As String
Dim NewRange As String
Dim LastCol As Long
Dim lastRow As Long
'Set Pivot Table & Source Worksheet
Set Data_Sheet = ThisWorkbook.Worksheets("PivotTableData3")
Set Pivot_Sheet = ThisWorkbook.Worksheets("Pivot3")
'Enter in Pivot Table Name
PivotName = "PivotTable2"
'Defining Staring Point & Dynamic Range
Data_Sheet.Activate
Set StartPoint = Data_Sheet.Range("A1")
LastCol = StartPoint.End(xlToRight).Column
DownCell = StartPoint.End(xlDown).Row
Set DataRange = Data_Sheet.Range(StartPoint, Cells(DownCell, LastCol))
NewRange = Data_Sheet.Name & "!" & DataRange.Address(ReferenceStyle:=xlR1C1)
'Change Pivot Table Data Source Range Address
Pivot_Sheet.PivotTables(PivotName). _
ChangePivotCache ActiveWorkbook. _
PivotCaches.Create(SourceType:=xlDatabase, SourceData:=NewRange)
'Ensure Pivot Table is Refreshed
Pivot_Sheet.PivotTables(PivotName).RefreshTable
'Complete Message
Pivot_Sheet.Activate
MsgBox "Your Pivot Table is now updated."
End Sub
'To enable GetPivotData function
Sub activateGetPivotData(control As IRibbonControl)
Application.GenerateGetPivotData = True
End Sub
'To disable GetPivotData function
Sub deactivateGetPivotData(control As IRibbonControl)
Application.GenerateGetPivotData = False
End Sub

