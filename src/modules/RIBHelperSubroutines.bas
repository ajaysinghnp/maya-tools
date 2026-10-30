Attribute VB_Name = "RIBHelperSubroutines"
Option Explicit

'===========================================================================
' RIBHelperSubroutines
'
' The work done by the buttons of the Helper group. These procedures are
' called from RibbonCallbacks (the single entry point of every ribbon
' button); they are not ribbon callbacks themselves.
'
' Rules followed here:
'  - never use Select or Activate to do the work, write to ranges directly
'  - Cancel (or Esc) in a dialog just stops, it never raises an error
'  - anything that can fail shows one clear message instead of a run-time error
'===========================================================================

Private Const TOC_SHEET_NAME As String = "Table of Content"

'---------------------------------------------------------------------------
' Table of content
'---------------------------------------------------------------------------

' Adds a "Table of Content" sheet as the first sheet of the ACTIVE workbook,
' with a link to every visible worksheet. An existing one is replaced.
Public Sub TableofContent()
    Dim wb As Workbook
    Dim toc As Worksheet
    Dim oldToc As Worksheet
    Dim ws As Worksheet
    Dim rowNo As Long
    Dim reason As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open a workbook first, then try again.", vbInformation, ADDIN_NAME
        Exit Sub
    End If
    If wb.ProtectStructure Then
        MsgBox "The structure of this workbook is protected, so a sheet cannot be added." & vbCrLf & _
               "Unprotect the workbook structure first.", vbExclamation, ADDIN_NAME
        Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False

    'Add the new sheet first: a workbook must always keep one visible sheet,
    'so the old table of content can then be deleted safely.
    Set toc = wb.Worksheets.Add(Before:=wb.Sheets(1))

    Set oldToc = FindWorksheet(wb, TOC_SHEET_NAME)
    If Not oldToc Is Nothing Then
        Application.DisplayAlerts = False
        oldToc.Delete
        Application.DisplayAlerts = True
    End If
    toc.Name = TOC_SHEET_NAME

    rowNo = 0
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If ws.Name <> toc.Name Then
                rowNo = rowNo + 1
                toc.Hyperlinks.Add Anchor:=toc.Cells(rowNo, 1), _
                                   Address:="", _
                                   SubAddress:="'" & Replace(ws.Name, "'", "''") & "'!A1", _
                                   ScreenTip:=ws.Name, _
                                   TextToDisplay:=ws.Name
            End If
        End If
    Next ws

    toc.Columns(1).AutoFit
    toc.Activate
    toc.Range("A1").Select
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    reason = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    'remove the half-made sheet, but never the finished table of content
    If Not toc Is Nothing Then
        If toc.Name <> TOC_SHEET_NAME Then toc.Delete
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    ReportError "create the table of content", reason
End Sub

'---------------------------------------------------------------------------
' Serial numbers and letters (written downwards from the active cell)
'---------------------------------------------------------------------------

Public Sub AddSerialNumbers()
    Dim n As Long
    Dim i As Long
    Dim values() As Variant

    If Not HasCellSelection() Then Exit Sub
    On Error GoTo Fail

    n = PromptForCount("How many serial numbers do you want to add?" & vbCrLf & _
                       "They are written downwards from the active cell.", RowsFromActiveCell())
    If n = 0 Then Exit Sub

    ReDim values(1 To n, 1 To 1)
    For i = 1 To n
        values(i, 1) = i
    Next i
    WriteDown values, n
    Exit Sub

Fail:
    ReportError "add the serial numbers", Err.Description
End Sub

Public Sub addcAlphabets()
    FillLetters 65, "capital letters"
End Sub

Public Sub addsAlphabets()
    FillLetters 97, "small letters"
End Sub

' firstCode is the character code of the first letter (65 = A, 97 = a).
Private Sub FillLetters(ByVal firstCode As Long, ByVal description As String)
    Dim n As Long
    Dim i As Long
    Dim room As Long
    Dim values() As Variant

    If Not HasCellSelection() Then Exit Sub
    On Error GoTo Fail

    room = RowsFromActiveCell()
    If room > 26 Then room = 26

    'The default is the whole alphabet, so pressing Enter adds A-Z (or a-z)
    n = PromptForCount("How many " & description & " do you want to add (1 to " & room & ")?" & vbCrLf & _
                       "They are written downwards from the active cell.", room, room)
    If n = 0 Then Exit Sub

    ReDim values(1 To n, 1 To 1)
    For i = 1 To n
        values(i, 1) = Chr$(firstCode + i - 1)
    Next i
    WriteDown values, n
    Exit Sub

Fail:
    ReportError "add the letters", Err.Description
End Sub

'---------------------------------------------------------------------------
' Insert columns and rows
'---------------------------------------------------------------------------

' Inserts the asked number of columns to the left of the active cell's column.
Public Sub InsertMultipleColumns()
    Dim n As Long

    If Not HasCellSelection() Then Exit Sub
    On Error GoTo Fail

    n = PromptForCount("How many columns do you want to insert?" & vbCrLf & _
                       "They are inserted to the left of the active cell's column.", _
                       ActiveCell.Worksheet.Columns.Count - ActiveCell.Column + 1)
    If n = 0 Then Exit Sub

    ActiveCell.Resize(1, n).EntireColumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Exit Sub

Fail:
    ReportError "insert the columns", Err.Description
End Sub

' Inserts the asked number of rows above the active cell's row.
Public Sub InsertMultipleRows()
    Dim n As Long

    If Not HasCellSelection() Then Exit Sub
    On Error GoTo Fail

    n = PromptForCount("How many rows do you want to insert?" & vbCrLf & _
                       "They are inserted above the active cell's row.", RowsFromActiveCell())
    If n = 0 Then Exit Sub

    ActiveCell.Resize(n, 1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Exit Sub

Fail:
    ReportError "insert the rows", Err.Description
End Sub

'---------------------------------------------------------------------------
' Private helpers
'---------------------------------------------------------------------------

' True when the active sheet is a worksheet, so there is an active cell.
' (With no workbook open, or a chart sheet active, there is no active cell.)
Private Function HasCellSelection() As Boolean
    If TypeName(ActiveSheet) = "Worksheet" Then
        HasCellSelection = True
    Else
        MsgBox "Select a cell on a worksheet first.", vbInformation, ADDIN_NAME
    End If
End Function

' Number of rows from the active cell to the bottom of the sheet.
Private Function RowsFromActiveCell() As Long
    RowsFromActiveCell = ActiveCell.Worksheet.Rows.Count - ActiveCell.Row + 1
End Function

' Asks for a whole number from 1 to maxValue. Asks again after an invalid
' answer. Returns 0 when the user cancels (Cancel button or Esc).
Private Function PromptForCount(ByVal message As String, ByVal maxValue As Long, _
                                Optional ByVal defaultValue As Long = 0) As Long
    Dim answer As Variant

    Do
        If defaultValue > 0 Then
            answer = Application.InputBox(message, ADDIN_NAME, defaultValue, Type:=1)
        Else
            answer = Application.InputBox(message, ADDIN_NAME, Type:=1)
        End If

        'Cancel returns the Boolean False
        If VarType(answer) = vbBoolean Then Exit Function

        If IsNumeric(answer) Then
            If answer = Int(answer) And answer >= 1 And answer <= maxValue Then
                PromptForCount = CLng(answer)
                Exit Function
            End If
        End If

        MsgBox "Enter a whole number from 1 to " & Format(maxValue, "#,##0") & ".", vbExclamation, ADDIN_NAME
    Loop
End Function

' Writes a one-column array into n cells downwards from the active cell and
' selects them. Asks first if any of those cells already holds something.
Private Sub WriteDown(ByRef values() As Variant, ByVal n As Long)
    Dim target As Range

    Set target = ActiveCell.Resize(n, 1)
    If Application.WorksheetFunction.CountA(target) > 0 Then
        If MsgBox("Some of the cells " & target.Address(False, False) & " already contain data." & vbCrLf & _
                  "Replace them?", vbYesNo + vbExclamation + vbDefaultButton2, ADDIN_NAME) = vbNo Then Exit Sub
    End If

    target.Value = values
    target.Select
End Sub

Private Function FindWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub ReportError(ByVal action As String, ByVal reason As String)
    MsgBox "MayaTools could not " & action & "." & vbCrLf & vbCrLf & reason, vbExclamation, ADDIN_NAME
End Sub
