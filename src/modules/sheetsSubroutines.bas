Attribute VB_Name = "sheetsSubroutines"
Sub ReverseSheets(control As IRibbonControl)
'Update 20140526
Dim xCount As Integer
xCount = Application.ActiveWorkbook.Worksheets.Count
For i = 1 To xCount - 1
    Application.Worksheets(1).Move After:=Application.Worksheets(xCount - i + 1)
Next
MsgBox ("Sheets order are reversed successfully!")
End Sub

Sub Splitdatabycol(control As IRibbonControl)
Dim lr As Long
Dim ws As Worksheet
Dim vcol, i As Integer
Dim icol As Long
Dim myarr As Variant
Dim title As String
Dim titlerow As Integer
Dim xTRg As Range
Dim xVRg As Range
Dim xWSTRg As Worksheet
On Error Resume Next
Set xTRg = Application.InputBox("Please select the header rows:", "MayaTools for Excel", "ss", Type:=8)
If TypeName(xTRg) = "Nothing" Then Exit Sub
Set xVRg = Application.InputBox("Please select the column you want to split data based on:", "MayaTools for Excel", "", Type:=8)
If TypeName(xVRg) = "Nothing" Then Exit Sub
vcol = xVRg.Column
Set ws = xTRg.Worksheet
lr = ws.Cells(ws.Rows.Count, vcol).End(xlUp).Row
title = xTRg.AddressLocal
titlerow = xTRg.Cells(1).Row
icol = ws.Columns.Count
ws.Cells(1, icol) = "Unique"
Application.DisplayAlerts = False
If Not Evaluate("=ISREF('xTRgWs_Sheet!A1')") Then
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "xTRgWs_Sheet"
Else
Sheets("xTRgWs_Sheet").Delete
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "xTRgWs_Sheet"
End If
Set xWSTRg = Sheets("xTRgWs_Sheet")
xTRg.Copy
xWSTRg.Paste Destination:=xWSTRg.Range("A1")
ws.Activate
For i = (titlerow + xTRg.Rows.Count) To lr
On Error Resume Next
If ws.Cells(i, vcol) <> "" And Application.WorksheetFunction.Match(ws.Cells(i, vcol), ws.Columns(icol), 0) = 0 Then
ws.Cells(ws.Rows.Count, icol).End(xlUp).Offset(1) = ws.Cells(i, vcol)
End If
Next
myarr = Application.WorksheetFunction.Transpose(ws.Columns(icol).SpecialCells(xlCellTypeConstants))
ws.Columns(icol).Clear
For i = 2 To UBound(myarr)
ws.Range(title).AutoFilter field:=vcol, Criteria1:=myarr(i) & ""
If Not Evaluate("=ISREF('" & myarr(i) & "'!A1)") Then
Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = myarr(i) & ""
Else
Sheets(myarr(i) & "").Move After:=Worksheets(Worksheets.Count)
End If
xWSTRg.Range(title).Copy
Sheets(myarr(i) & "").Paste Destination:=Sheets(myarr(i) & "").Range("A1")
ws.Range("A" & (titlerow + xTRg.Rows.Count) & ":A" & lr).EntireRow.Copy Sheets(myarr(i) & "").Range("A" & (titlerow + xTRg.Rows.Count))
Sheets(myarr(i) & "").Columns.AutoFit
Next
xWSTRg.Delete
ws.AutoFilterMode = False
ws.Activate
Application.DisplayAlerts = True
End Sub


Sub HideWorksheet(control As IRibbonControl)
'hide all the worksheets in your workbook other than the active worksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
ws.Visible = xlSheetHidden
End If
Next ws
End Sub
Sub UnhideAllWorksheet(control As IRibbonControl)
'un-hide all the worksheets which you have hide with previous code
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Visible = xlSheetVisible
Next ws
End Sub
Sub DeleteWorksheets(control As IRibbonControl)
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If ws.Name <> ThisWorkbook.ActiveSheet.Name Then
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
End If
Next ws
End Sub

Sub ProtectAllWorskeets(control As IRibbonControl)
Dim ws As Worksheet
Dim ps As String
ps = InputBox("Enter a Password.", vbOKCancel)
For Each ws In ActiveWorkbook.Worksheets
ws.Protect Password:=ps
Next ws
End Sub

Sub Resize_Charts(control As IRibbonControl)
Dim i As Integer
For i = 1 To ActiveSheet.ChartObjects.Count
With ActiveSheet.ChartObjects(i)
.Width = 300
.Height = 200
End With
Next i
End Sub

Sub InsertMultipleSheets(control As IRibbonControl)
Dim i As Integer
i = InputBox("Enter number of sheets to insert.", "Enter Multiple Sheets")
Sheets.Add After:=ActiveSheet, Count:=i
End Sub

Sub ProtectWS(control As IRibbonControl)
ActiveSheet.Protect "mypassword", True, True
End Sub

Sub UnprotectWS(control As IRibbonControl)
ActiveSheet.Unprotect "mypassword"
End Sub

Sub SortWorksheets(control As IRibbonControl)
Dim i As Integer
Dim j As Integer
Dim iAnswer As VbMsgBoxResult
iAnswer = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
& "Clicking No will sort in Descending Order", _
vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
For i = 1 To Sheets.Count
For j = 1 To Sheets.Count - 1
If iAnswer = vbYes Then
If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
Sheets(j).Move After:=Sheets(j + 1)
End If
ElseIf iAnswer = vbNo Then
If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then Sheets(j).Move After:=Sheets(j + 1)
End If
'End If
Next j
Next i
End Sub

Sub lockCellsWithFormulas(control As IRibbonControl)
'To protect cell with formula with a single click you can use this code
With ActiveSheet
.Unprotect
.Cells.Locked = False
.Cells.SpecialCells(xlCellTypeFormulas).Locked = True
.Protect AllowDeletingRows:=True
End With
End Sub

Sub deleteBlankWorksheets(control As IRibbonControl)
'it will check all the worksheets in the active workbook and delete if a worksheet is blank.
Dim ws As Worksheet
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each ws In Application.Worksheets
If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
ws.Delete
End If
Next
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub UnhideRowsColumns(control As IRibbonControl)
Columns.EntireColumn.Hidden = False
Rows.EntireRow.Hidden = False
End Sub

Sub SaveWorkshetAsPDF(control As IRibbonControl)
    Dim ws As Worksheet
    Dim folderPath As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Select folder to save PDFs"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1) & "\"
    End With

    For Each ws In ActiveWorkbook.Worksheets
        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=folderPath & ws.Name & ".pdf"
    Next ws
End Sub


Sub DisablePageBreaks(control As IRibbonControl)
Dim wb As Workbook
Dim wks As Worksheet
Application.ScreenUpdating = False
For Each wb In Application.Workbooks
'For Each Sht In wb.WorksheetsSht.DisplayPageBreaks= False
Next Sht
Next wb
Application.ScreenUpdating = True
End Sub
