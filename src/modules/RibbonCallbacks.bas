Attribute VB_Name = "RibbonCallbacks"
Option Explicit

'===========================================================================
' RibbonCallbacks
'
' The ONE place where every ribbon button enters the add-in.
'
'  - ribbon/customUI14.xml and ribbon/customUI.xml point every onAction here
'    (Ribbon_<Action>). No other module may be a ribbon callback.
'  - A callback does as little as possible: check that it can run, then call
'    the procedure that does the work. Work that is only a line or two
'    (About, Website) is written here directly.
'  - To add a button: add the button to both ribbon XML files with
'    onAction="Ribbon_<Action>", then add its callback in the matching
'    section below.
'
' Callback signature (required by Excel): Sub Name(control As IRibbonControl)
'===========================================================================

Private Const WEBSITE_URL As String = "https://ajaysingh.com.np/"

'---------------------------------------------------------------------------
' Shared helpers
'---------------------------------------------------------------------------

' ActiveWorkbook is Nothing when no workbook is open, and using it then raises
' run-time error 91. Buttons that work on a workbook call this first.
Private Function RequireWorkbook() As Boolean
    RequireWorkbook = Not (ActiveWorkbook Is Nothing)
    If Not RequireWorkbook Then
        MsgBox "Open a workbook first, then try again.", vbInformation, ADDIN_NAME
    End If
End Function

'---------------------------------------------------------------------------
' Sheets group
'---------------------------------------------------------------------------

'Callback for split_sheets_by_column onAction
Public Sub Ribbon_SplitSheetsByColumn(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    Splitdatabycol control
End Sub

'Callback for insert_multiple_sheets onAction
Public Sub Ribbon_InsertSheets(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    InsertMultipleSheets control
End Sub

'Callback for sort_sheets onAction
Public Sub Ribbon_SortSheets(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    SortWorksheets control
End Sub

'Callback for reverse_sheets_order onAction
Public Sub Ribbon_ReverseSheets(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    ReverseSheets control
End Sub

'---------------------------------------------------------------------------
' Workbooks group
'---------------------------------------------------------------------------

'Callback for file_backup onAction
Public Sub Ribbon_BackUp(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    FileBackUp control
End Sub

'Callback for Copy_Worksheet_To_New_Workbook onAction
Public Sub Ribbon_SheetToNewBook(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    CopyWorksheetToNewWorkbook control
End Sub

'Callback for visible_workbooks onAction
Public Sub Ribbon_VisibleSheets(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    VisibleWorkbooks control
End Sub

'Callback for close_all_workbooks onAction
Public Sub Ribbon_CloseAllWorkbooks(control As IRibbonControl)
    If ActiveWorkbook Is Nothing Then
        MsgBox "There are no open workbooks to close.", vbInformation, ADDIN_NAME
        Exit Sub
    End If
    CloseAllWorkbooks control
End Sub

'---------------------------------------------------------------------------
' Formatting group
'---------------------------------------------------------------------------

'Callback for Highlight_Duplicate_Values onAction
Public Sub Ribbon_HighlightDuplicates(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    HighlightDuplicateValues control
End Sub

'Callback for Highlight_Greater_Than_Values onAction
Public Sub Ribbon_HighlightGreater(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    HighlightGreaterThanValues control
End Sub

'Callback for Highlight_Lower_ThanValues onAction
Public Sub Ribbon_HighlightLesser(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    HighlightLowerThanValues control
End Sub

'Callback for highlight_Negative_Numbers onAction
Public Sub Ribbon_HighlightNegatives(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    highlightNegativeNumbers control
End Sub

'---------------------------------------------------------------------------
' Tax Helper group
'---------------------------------------------------------------------------

'Callback for fill_formatting onAction
Public Sub Ribbon_FillFormattingSummary(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    fillFormat control
End Sub

'Callback for text_to_numbers onAction
Public Sub Ribbon_ExtractNumber(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    ExtrNumbersFromRange control
End Sub

'---------------------------------------------------------------------------
' Helper group
'---------------------------------------------------------------------------

'Callback for table_of_content onAction
Public Sub Ribbon_TableOfContent(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    TableofContent
End Sub

'Callback for add_serial_numbers onAction
Public Sub Ribbon_SerialNumbers(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    AddSerialNumbers
End Sub

'Callback for add_capital_alphabets onAction
Public Sub Ribbon_CapitalAlphabets(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    addcAlphabets
End Sub

'Callback for add_small_alphabets onAction
Public Sub Ribbon_SmallAlphabets(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    addsAlphabets
End Sub

'Callback for insert_multiple_columns onAction
Public Sub Ribbon_InsertColumns(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    InsertMultipleColumns
End Sub

'Callback for insert_multiple_rows onAction
Public Sub Ribbon_InsertRows(control As IRibbonControl)
    If Not RequireWorkbook() Then Exit Sub
    InsertMultipleRows
End Sub

'---------------------------------------------------------------------------
' Developer group (these do not need an open workbook)
'---------------------------------------------------------------------------

'Callback for about_developer onAction
Public Sub Ribbon_AboutDeveloper(control As IRibbonControl)
    aboutDeveloper.show
End Sub

'Callback for web_link onAction
' Uses ThisWorkbook (the add-in itself, always loaded), not ActiveWorkbook,
' so it also works when no workbook is open.
Public Sub Ribbon_WebLink(control As IRibbonControl)
    ThisWorkbook.FollowHyperlink Address:=WEBSITE_URL, NewWindow:=True
End Sub
