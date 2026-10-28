Attribute VB_Name = "workBookSubroutines"
Sub FileBackUp(control As IRibbonControl)
ThisWorkbook.SaveCopyAs Filename:=ThisWorkbook.Path & _
"" & Format(Date, "mm-dd-yy") & " " & _
ThisWorkbook.Name
End Sub

Sub CloseAllWorkbooks(control As IRibbonControl)
Dim wbs As Workbook
For Each wbs In Workbooks
wbs.Close SaveChanges:=True
Next wb
End Sub

Sub CopyWorksheetToNewWorkbook(control As IRibbonControl)
ThisWorkbook.ActiveSheet.Copy Before:=Workbooks.Add.Worksheets(1)
End Sub

'You can change the subject, email, and body text in code and if you want to send this mail directly, use ".Send" instead of ".Display".
Sub Send_Mail(control As IRibbonControl)
Dim OutApp As Object
Dim OutMail As Object
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
With OutMail
.To = "Sales@FrontLinePaper.com"
.Subject = "Growth Report"
.Body = "Hello Team, Please find attached Growth Report."
.Attachments.Add ActiveWorkbook.FullName
.display
End With
Set OutMail = Nothing
Set OutApp = Nothing
End Sub
'it will open your default mail client and attached active workbook with it as an attachment.
Sub OpenWorkbookAsAttachment(control As IRibbonControl)
Application.Dialogs(xlDialogSendMail).Show
End Sub

Sub auto_open1()
MsgBox "Welcome To MayaTools for Excel. Find the MayaTools in menu to enjoy!"
End Sub

Sub auto_close1()
MsgBox "Bye Bye! Don't forget to check other cool stuff on https://ajaysingh.com.np/"""
End Sub
'get the number of workbooks which are not saved yet.
Sub VisibleWorkbooks(control As IRibbonControl)
Dim book As Workbook
Dim i As Integer
For Each book In Workbooks
If book.Saved = False Then
i = i + 1
End If
Next book
MsgBox i
End Sub
