Attribute VB_Name = "RIBTaxHelperSubroutines"

Attribute VB_Name = "TaxCalculationFormat"
Sub fillFormat(control As IRibbonControl)
'Subroutine to fill the format by calling sumFormat function

Dim cell, credits, balance, temp As Range
Dim tax_on_interest As Double
Dim completed_color, incomplete_color As Integer
Dim values_color, header_color, danger_color, danger_bg_color, success_bg_color, success_color As Long
Dim total_transactions, entered_transactions, remaining_transactions As Integer
Dim not_okay, okay As FormatCondition

'Setting the reference to the active cell
Set cell = Application.ActiveCell
cell.Value = "Initializing..."

'setting colors values
values_color = RGB(102, 255, 153)
header_color = RGB(196, 196, 196)
danger_color = RGB(156, 0, 6)
success_color = RGB(0, 97, 0)
success_bg_color = RGB(198, 239, 206)
danger_bg_color = RGB(255, 199, 206)

'Getting the values for the population
cell.Value = "Gathering informations..."
Set credits = Application.InputBox(prompt:="Select the credit range:", title:="Credit List", Type:=8)
Set balance = Application.InputBox(prompt:="Select Accumulated Balance:", title:="Accumulated Balance", Type:=8)
tax_on_interest = Application.InputBox("Tax On Interest for this month:", "Tax On Interest", 0)
completed_color = cell.Offset(1, 2).Interior.ColorIndex
incomplete_color = cell.Offset(1, 6).Interior.ColorIndex

'Starting the population of data
cell.Value = "Generating..."
Range(cell, cell.Offset(0, 3)).UnMerge
Range(cell, cell.Offset(4, 7)).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

'Calculating data
cell.Value = "Calculating values and formulas"

'working on left column headers
cell.Offset(1, 0).Value = "Completed Color"
cell.Offset(1, 0).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(2, 0).Value = "Total Transactions"
cell.Offset(2, 0).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(3, 0).Value = "Transactions Entered"
cell.Offset(3, 0).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(4, 0).Value = "Transactions Remaining"
cell.Offset(4, 0).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
Range(cell.Offset(1, 0), cell.Offset(4, 0)).Interior.Color = header_color

'working on equals to
cell.Offset(1, 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(2, 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(2, 1).Value = "="
cell.Offset(3, 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(3, 1).Value = "="
cell.Offset(4, 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(4, 1).Value = "="

'working on left values
cell.Offset(1, 2).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(2, 2).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(3, 2).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(4, 2).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

Range(cell.Offset(1, 1), cell.Offset(1, 2)).Merge
Range(cell.Offset(1, 1), cell.Offset(1, 2)).HorizontalAlignment = xlCenter
cell.Offset(2, 2).Formula = "=COUNT(" & credits.Address & ") + 1"
cell.Offset(3, 2).Formula = "=CountCellsByColor(" & credits.Address & ", " & cell.Offset(1, 2).Address & ") + CountCellsByColor(" & cell.Offset(0, 6).Address & ", " & cell.Offset(1, 2).Address & ")"
cell.Offset(4, 2).Formula = "=" & cell.Offset(2, 2).Address & " - " & cell.Offset(3, 2).Address

Range(cell.Offset(2, 2), cell.Offset(4, 2)).Interior.Color = values_color

'working on right column headers
cell.Offset(1, 3).Value = "Incomplete Color"
cell.Offset(1, 3).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(3, 3).Value = "Sum of Entered"
cell.Offset(3, 3).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(2, 3).Value = "Total Sum"
cell.Offset(2, 3).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(4, 3).Value = "Difference"
cell.Offset(4, 3).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
Range(cell.Offset(1, 3), cell.Offset(4, 3)).Interior.Color = header_color

'working on equals to
cell.Offset(1, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(2, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(2, 4).Value = "="
cell.Offset(3, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(3, 4).Value = "="
cell.Offset(4, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(4, 4).Value = "="

'working on right values
cell.Offset(1, 5).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(2, 5).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(3, 5).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(4, 5).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

Range(cell.Offset(1, 4), cell.Offset(1, 5)).Merge
Range(cell.Offset(1, 4), cell.Offset(1, 5)).HorizontalAlignment = xlCenter
cell.Offset(1, 4).Value = "(No Color)"
cell.Offset(3, 5).Formula = "=SumCellsByColor(" & credits.Address & ", " & cell.Offset(1, 2).Address & ") + SumCellsByColor(" & cell.Offset(0, 5).Address & ", " & cell.Offset(1, 2).Address & ")"
cell.Offset(2, 5).Formula = "=SUM(" & credits.Address & ") + " & cell.Offset(0, 5).Address
cell.Offset(4, 5).Formula = "=" & cell.Offset(2, 5).Address & " - " & cell.Offset(3, 5).Address

Range(cell.Offset(2, 5), cell.Offset(4, 5)).Interior.Color = values_color
Range(cell.Offset(2, 5), cell.Offset(4, 5)).Style = "Currency"

'merging initial row
Range(cell, cell.Offset(0, 2)).Merge

'Working on Tax On Interest
Range(cell.Offset(0, 1), cell.Offset(0, 2)).Merge
Range(cell.Offset(0, 1), cell.Offset(0, 2)).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
Range(cell.Offset(0, 1), cell.Offset(0, 2)).Interior.Color = header_color
Range(cell.Offset(0, 1), cell.Offset(0, 2)).HorizontalAlignment = xlCenter
cell.Offset(0, 1).Value = "Tax On Interest"
cell.Offset(0, 3).Value = tax_on_interest
cell.Offset(0, 3).Style = "Currency"
cell.Offset(0, 3).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

'working on sum of interest
Range(cell.Offset(0, 4), cell.Offset(0, 5)).Merge
cell.Offset(0, 4).Value = "Status"
cell.Offset(0, 4).Interior.Color = header_color
cell.Offset(0, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(0, 4).HorizontalAlignment = xlCenter

cell.Offset(1, 4).Value = "Credit Sum"
cell.Offset(1, 4).Interior.Color = header_color
cell.Offset(1, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

cell.Offset(2, 4).Value = "Balance"
cell.Offset(2, 4).Interior.Color = header_color
cell.Offset(2, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

cell.Offset(3, 4).Value = "Difference"
cell.Offset(3, 4).Interior.Color = header_color
cell.Offset(3, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

cell.Offset(4, 4).Value = "Result"
cell.Offset(4, 4).Interior.Color = header_color
cell.Offset(4, 4).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

cell.Offset(1, 5).Formula = "=SUM(" & credits.Address & ")"
cell.Offset(1, 5).Style = "Currency"
cell.Offset(1, 5).Interior.Color = values_color
cell.Offset(1, 5).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

cell.Offset(2, 5).Formula = "=" & balance.Address
cell.Offset(2, 5).Style = "Currency"
cell.Offset(2, 5).Interior.Color = values_color
cell.Offset(2, 5).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

cell.Offset(3, 5).Formula = "=" & cell.Offset(1, 5).Address & " - " & cell.Offset(2, 5).Address
cell.Offset(3, 5).Style = "Currency"
cell.Offset(3, 5).Interior.Color = values_color
cell.Offset(3, 5).BorderAround LineStyle:=xlContinuous, Weight:=xlThin

cell.Offset(4, 5).Formula = "=IF(" & cell.Offset(3, 5).Address & " = 0 ,""Okay"",""Not Okay"")"
cell.Offset(4, 5).HorizontalAlignment = xlCenter
cell.Offset(4, 5).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
cell.Offset(4, 5).Font.Name = "Times New Roman"
cell.Offset(4, 5).Font.FontStyle = "Bold Italic"
cell.Offset(4, 5).FormatConditions.Delete

Set not_okay = cell.Offset(4, 5).FormatConditions.Add(xlTextString, TextOperator:=xlContains, String:="Not Okay")
Set okay = cell.Offset(4, 5).FormatConditions.Add(xlTextString, TextOperator:=xlContains, String:="Okay")

With okay
    .Font.Color = success_color
    .Interior.Color = success_bg_color
End With

With not_okay
    .Font.Color = danger_color
    .Interior.Color = danger_bg_color
End With

'setting completed and incompleted colors
cell.Value = "Setting colors"
MsgBox "Set your Completed and Incomplete transactions color"

'clearing heading
cell.Value = ""
End Sub
