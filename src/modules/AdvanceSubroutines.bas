Attribute VB_Name = "AdvanceSubroutines"

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

