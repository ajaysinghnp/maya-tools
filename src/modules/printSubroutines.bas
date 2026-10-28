Attribute VB_Name = "printSubroutines"
Sub printComments(control As IRibbonControl)
With ActiveSheet.PageSetup
.printComments = xlPrintSheetEnd
End With
End Sub

Sub printNarrowMargin(control As IRibbonControl)
With ActiveSheet.PageSetup
.LeftMargin = Application
.InchesToPoints (0.25)
.RightMargin = Application.InchesToPoints(0.25)
.TopMargin = Application.InchesToPoints(0.75)
.BottomMargin = Application.InchesToPoints(0.75)
.HeaderMargin = Application.InchesToPoints(0.3)
.FooterMargin = Application.InchesToPoints(0.3)
End With
'ActiveWindow.SelectedSheets.PrintOutCopies:=1, Collate:=True,
'IgnorePrintAreas:=False
End Sub

Sub printSelection(control As IRibbonControl)
'Selection.PrintOutCopies:=1, Collate:=True
End Sub

Sub printCustomSelection(control As IRibbonControl)
Dim startpage As Integer
Dim endpage As Integer
startpage = InputBox("Please Enter Start Page number.", "Enter Value")
If Not WorksheetFunction.IsNumber(startpage) Then
MsgBox "Invalid Start Page number. Please try again.", "Error"
Exit Sub
End If
endpage = InputBox("Please Enter End Page number.", "Enter Value")
If Not WorksheetFunction.IsNumber(endpage) Then
MsgBox "Invalid End Page number. Please try again.", "Error"
Exit Sub
End If
'Selection.PrintOutFrom:=startpage, To:=endpage, Copies:=1, Collate:=True
End Sub

