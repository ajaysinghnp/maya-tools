Attribute VB_Name = "HeaderFooterSubroutines"
Sub DateInHeader(control As IRibbonControl)
With ActiveSheet.PageSetup
.LeftHeader = ""
.CenterHeader = "&D"
.RightHeader = ""
.LeftFooter = ""
.CenterFooter = ""
.RightFooter = ""
End With
End Sub

Sub CustomHeader(control As IRibbonControl)
Dim myText As String
myText = InputBox("Enter your text here", "MayaTools for Excel")
With ActiveSheet.PageSetup
.LeftHeader = ""
.CenterHeader = myText
.RightHeader = ""
.LeftFooter = ""
.CenterFooter = ""
.RightFooter = ""
End With
End Sub


