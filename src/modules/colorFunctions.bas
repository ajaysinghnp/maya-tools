Attribute VB_Name = "colorFunctions"
Function GetCellColor(xlRange As Range)
    Dim indRow, indColumn As Long
    Dim arResults()
 
    Application.Volatile
 
    If xlRange Is Nothing Then
        Set xlRange = Application.ThisCell
    End If
 
    If xlRange.Count > 1 Then
      ReDim arResults(1 To xlRange.Rows.Count, 1 To xlRange.Columns.Count)
       For indRow = 1 To xlRange.Rows.Count
         For indColumn = 1 To xlRange.Columns.Count
           arResults(indRow, indColumn) = xlRange(indRow, indColumn).Interior.Color
         Next
       Next
     GetCellColor = arResults
    Else
     GetCellColor = xlRange.Interior.Color
    End If
End Function
 
Function GetCellFontColor(xlRange As Range)
    Dim indRow, indColumn As Long
    Dim arResults()
 
    Application.Volatile
 
    If xlRange Is Nothing Then
        Set xlRange = Application.ThisCell
    End If
 
    If xlRange.Count > 1 Then
      ReDim arResults(1 To xlRange.Rows.Count, 1 To xlRange.Columns.Count)
       For indRow = 1 To xlRange.Rows.Count
         For indColumn = 1 To xlRange.Columns.Count
           arResults(indRow, indColumn) = xlRange(indRow, indColumn).Font.Color
         Next
       Next
     GetCellFontColor = arResults
    Else
     GetCellFontColor = xlRange.Font.Color
    End If
 
End Function
 
Function CountCellsByColor(rData As Range, cellRefColor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cntRes As Long
 
    Application.Volatile
    cntRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Interior.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Interior.Color Then
            cntRes = cntRes + 1
        End If
    Next cellCurrent
 
    CountCellsByColor = cntRes
End Function
 
Function SumCellsByColor(rData As Range, cellRefColor As Range)
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim sumRes
 
    Application.Volatile
    sumRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Interior.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Interior.Color Then
            sumRes = WorksheetFunction.Sum(cellCurrent, sumRes)
        End If
    Next cellCurrent
 
    SumCellsByColor = sumRes
End Function
 
Function CountCellsByFontColor(rData As Range, cellRefColor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cntRes As Long
 
    Application.Volatile
    cntRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Font.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Font.Color Then
            cntRes = cntRes + 1
        End If
    Next cellCurrent
 
    CountCellsByFontColor = cntRes
End Function
 
Function SumCellsByFontColor(rData As Range, cellRefColor As Range)
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim sumRes
 
    Application.Volatile
    sumRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Font.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Font.Color Then
            sumRes = WorksheetFunction.Sum(cellCurrent, sumRes)
        End If
    Next cellCurrent
 
    SumCellsByFontColor = sumRes
End Function

