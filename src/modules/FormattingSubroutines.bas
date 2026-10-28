Attribute VB_Name = "FormattingSubroutines"
Sub HighlightDuplicateValues(control As IRibbonControl)
'Highlights duplicate values in the selected range
Dim myRange As Range
Dim myCell As Range
Set myRange = Selection
For Each myCell In myRange
If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
myCell.Interior.ColorIndex = 36
End If
Next myCell
End Sub

Sub TopTen(control As IRibbonControl)
'Highlight Top Ten Values from the selected range
Selection.FormatConditions.AddTop10
Selection.FormatConditions(Selection.FormatConditions.Count).S
tFirstPriority
With Selection.FormatConditions(1)
.TopBottom = xlTop10Top
.Rank = 10
.Percent = False
End With
With Selection.FormatConditions(1).Font
.Color = -16752384
.TintAndShade = 0
End With
With Selection.FormatConditions(1).Interior
.PatternColorIndex = xlAutomatic
.Color = 13561798
.TintAndShade = 0
End With
Selection.FormatConditions(1).StopIfTrue = False
End Sub

Sub HighlightRanges(control As IRibbonControl)
Dim RangeName As Name
Dim HighlightRange As Range
On Error Resume Next
For Each RangeName In ActiveWorkbook.Names
Set HighlightRange = RangeName.RefersToRange
HighlightRange.Interior.ColorIndex = 36
Next RangeName
End Sub

Sub HighlightGreaterThanValues(control As IRibbonControl)
Dim i As Integer
i = InputBox("Enter Greater Than Value", "MayaTools for Excel")
Selection.FormatConditions.Delete
Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula:=i
Selection.FormatConditions(Selection.FormatConditions.Count).StFirstPriority
With Selection.FormatConditions(1)
.Font.Color = RGB(0, 0, 0)
.Interior.Color = RGB(31, 218, 154)
End With
End Sub

Sub HighlightLowerThanValues(control As IRibbonControl)
Dim i As Integer
i = InputBox("Enter Lower Than Value", "MayaTools for Excel")
Selection.FormatConditions.Delete
Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLower, Formula1:=i
Selection.FormatConditions(Selection.FormatConditions.Count).StFirstPriority
With Selection.FormatConditions(1)
.Font.Color = RGB(0, 0, 0)
.Interior.Color = RGB(217, 83, 79)
End With
End Sub

Sub highlightNegativeNumbers(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection
If WorksheetFunction.IsNumber(rng) Then
If rng.Value < 0 Then
rng.Font.Color = -16776961
End If
End If
Next
End Sub

Sub highlightValue(control As IRibbonControl)
Dim myStr As String
Dim myRg As Range
Dim myTxt As String
Dim myCell As Range
Dim myChar As String
Dim i As Long
Dim j As Long
On Error Resume Next
If ActiveWindow.RangeSelection.Count > 1 Then
myTxt = ActiveWindow.RangeSelection.AddressLocal
Else
myTxt = ActiveSheet.UsedRange.AddressLocal
End If
LInput: Set myRg = Application.InputBox("Please select the Data range:", "MayaTools for Excel", myTxt, , , , , 8)
If myRg Is Nothing Then
Exit Sub
End If
If myRg.Areas.Count > 1 Then
MsgBox "Doesnot support multiple columns", , "MayaTools for Excel"
GoTo LInput
End If
If myRg.Columns.Count <> 2 Then
MsgBox "the selected range can only contain two columns "
GoTo LInput
End If
For i = 0 To myRg.Rows.Count - 1
myStr = myRg.Range("B1").Offset(i, 0).Value
With myRg.Range("A1").Offset(i, 0)
.Font.ColorIndex = 1
For j = 1 To Len(.Text)
Mid(Text, j, Len(myStr)) = myStrThen
.Characters(j, Len(myStr)).Font.ColorIndex = 3
Next
End With
Next i
End Sub

Sub highlightCommentCells(control As IRibbonControl)
Selection.SpecialCells(xlCellTypeComments).Select
Selection.Style = "Note"
End Sub

Sub highlightAlternateRows(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection.Rows
If rng.Row Mod 2 = 1 Then
rng.Style = "20% -Accent1"
rng.Value = rng ^ (1 / 3)
Else
End If
Next rng
End Sub

Sub HighlightMisspelledCells(control As IRibbonControl)
Dim rng As Range
For Each rng In ActiveSheet.UsedRange
If Not Application.CheckSpelling(word:=rng.Text) Then
rng.Style = "Bad"
End If
Next rng
End Sub

Sub highlightErrors(control As IRibbonControl)
Dim rng As Range
Dim i As Integer
For Each rng In ActiveSheet.UsedRange
If WorksheetFunction.IsError(rng) Then
i = i + 1
rng.Style = "Bad"
End If
Next rng
MsgBox "There are total " & i & " error(s) in this worksheet."
End Sub

Sub highlightSpecificValues(control As IRibbonControl)
'Highlight Cells with a Specific Text in Worksheet
Dim rng As Range
Dim i As Integer
Dim c As Variant
c = InputBox("Enter Value To Highlight")
For Each rng In ActiveSheet.UsedRange
If rng = c Then
rng.Style = "Note"
i = i + 1
End If
Next rng
MsgBox "There are total " & i & " " & c & " in this worksheet."
End Sub

Sub blankWithSpace(control As IRibbonControl)
'Highlight all the Blank Cells Invisible Space
Dim rng As Range
For Each rng In ActiveSheet.UsedRange
If rng.Value = " " Then
rng.Style = "Note"
End If
Next rng
End Sub

Sub highlightMaxValue(control As IRibbonControl)
'Highlight Max Value In The Range
Dim rng As Range
For Each rng In Selection
If rng = WorksheetFunction.Max(Selection) Then
rng.Style = "Good"
End If
Next rng
End Sub

Sub highlightMinValue(control As IRibbonControl)
'Highlight Min Value In The Range
Dim rng As Range
For Each rng In Selection
If rng = WorksheetFunction.Min(Selection) Then
rng.Style = "Good"
End If
Next rng
End Sub

Sub highlightUniqueValues(control As IRibbonControl)
'Highlight Unique Values
Dim rng As Range
Set rng = Selection
rng.FormatConditions.Delete
Dim uv As UniqueValues
Set uv = rng.FormatConditions.AddUniqueValues
uv.DupeUnique = xlUnique
uv.Interior.Color = vbGreen
End Sub

Sub columnDifference(control As IRibbonControl)
'Using this code you can highlight the difference between two columns (corresponding cells)
Range("H7:H8,I7:I8").Select
Selection.ColumnDifferences(ActiveCell).Select
Selection.Style = "Bad"
End Sub

Sub rowDifference(control As IRibbonControl)
'And by using this code you can highlight difference between two row (corresponding cells)
Range("H7:H8,I7:I8").Select
Selection.RowDifferences(ActiveCell).Select
Selection.Style = "Bad"
End Sub

Sub convertUpperCase(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection
If Application.WorksheetFunction.IsText(rng) Then
rng.Value = UCase(rng)
End If
Next
End Sub
Sub convertLowerCase(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection
If Application.WorksheetFunction.IsText(rng) Then
rng.Value = LCase(rng)
End If
Next
End Sub
Sub convertProperCase(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection
If WorksheetFunction.IsText(rng) Then
rng.Value = WorksheetFunction.Proper(rng.Value)
End If
Next
End Sub

Sub convertTextCase(control As IRibbonControl)
Dim rng As Range
For Each rng In Selection
If WorksheetFunction.IsText(rng) Then
rng.Value = UCase(Left(rng, 1)) & LCase(Right(rng, Len(rng) - 1))
End If
Next rng
End Sub

Sub removeChar(control As IRibbonControl)
Dim rng As Range
Dim rc As String
rc = InputBox("Character(s) to Replace", "Enter Value")
For Each rng In Selection
Selection.Replace What:=rc, Replacement:=""
Next
End Sub
Sub Word_Count_Worksheet(control As IRibbonControl)
Dim WordCnt As Long
Dim rng As Range
Dim S As String
Dim N As Long
For Each rng In ActiveSheet.UsedRange.Cells
S = Application.WorksheetFunction.Trim(rng.Text)
N = 0
If S <> vbNullString Then
N = Len(S) - Len(Replace(S, " ", "")) + 1
End If
WordCnt = WordCnt + N
Next rng
MsgBox "There are total " & Format(WordCnt, "#,##0") & " words in the active worksheet"
End Sub
Sub removeApostrophes(control As IRibbonControl)
Selection.Value = Selection.Value
End Sub
Sub removeDecimals(control As IRibbonControl)
Dim lnumber As Double
Dim lResult As Long
Dim rng As Range
For Each rng In Selection
rng.Value = Int(rng)
rng.NumberFormat = "0"
Next rng
End Sub
Sub multiplyWithNumber(control As IRibbonControl)
Dim rng As Range
Dim c As Integer
c = InputBox("Enter number to multiple", "Input Required")
For Each rng In Selection
If WorksheetFunction.IsNumber(rng) Then
rng.Value = rng * c
Else
End If
Next rng
End Sub
Sub addNumber(control As IRibbonControl)
Dim rng As Range
Dim i As Integer
i = InputBox("Enter number to multiple", "Input Required")
For Each rng In Selection
If WorksheetFunction.IsNumber(rng) Then
rng.Value = rng + i
Else
End If
Next rng
End Sub
