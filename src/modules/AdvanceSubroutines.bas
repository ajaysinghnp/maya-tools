Attribute VB_Name = "AdvanceSubroutines"
'Callback for about_developer onAction
Sub aboutDev(control As IRibbonControl)
aboutDeveloper.Show
End Sub

'Callback for web_link onAction
Sub webLink(control As IRibbonControl)
    'Dim sXLFile As String
    'Dim sFolder As String
    Dim sWebsite As String
   
    'sFolder = "C:\Temp\" ' You can change as per your requirement
    'sXLFile = "C:\Temp\test1.xls" ' You can change as per your requirement
    sWebsite = "https://ajaysingh.com.np/" ' You can change as per your requirement

    'ActiveWorkbook.FollowHyperlink Address:=sFolder, NewWindow:=True 'Open Folder
    'ActiveWorkbook.FollowHyperlink Address:=sXLFile, NewWindow:=True 'Open excel workbook
    ActiveWorkbook.FollowHyperlink Address:=sWebsite, NewWindow:=True 'Open Website
End Sub

Sub TableofContent(control As IRibbonControl)
Dim i As Long
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("Table of Content").Delete
Application.DisplayAlerts = True
On Error GoTo 0
ThisWorkbook.Sheets.Add Before:=ThisWorkbook.Worksheets(1)
ActiveSheet.Name = "Table of Content"
For i = 1 To Sheets.Count
With ActiveSheet
.Hyperlinks.Add _
Anchor:=ActiveSheet.Cells(i, 1), _
Address:="", _
SubAddress:="'" & Sheets(i).Name & "'!A1", _
ScreenTip:=Sheets(i).Name, _
TextToDisplay:=Sheets(i).Name
End With
Next i
End Sub


Sub PasteAsPicture(control As IRibbonControl)
Application.CutCopyMode = False
Selection.Copy
ActiveSheet.Pictures.Paste.Select
End Sub

Sub LinkedPicture(control As IRibbonControl)
Selection.Copy
ActiveSheet.Pictures.Paste(Link:=True).Select
End Sub

Sub Speak(control As IRibbonControl)
Selection.Speak
End Sub

Sub SearchWindow32(control As IRibbonControl)
Dim chromePath As String
Dim search_string As String
Dim query As String
query = InputBox("Enter here your search here", "Google Search")
search_string = query
search_string = Replace(search_string, " ", "+")
'Uncomment the following line for Windows 64 versions and comment out Windows 32 versions'
chromePath = "C:ProgramFilesGoogleChromeApplicationchrome.exe """
'Uncomment the following line for Windows 32 versions and comment out Windows 64 versions
chromePath = "C:Program Files(x86)GoogleChromeApplicationchrome.exe"
Shell (chromePath & " -url http://google.com/#q=" & search_string)
End Sub

