Attribute VB_Name = "ChartSubroutines"
Sub ChangeChartType(control As IRibbonControl)
ActiveChart.ChartType = xlColumnClustered
End Sub

Sub ConvertChartToPicture(control As IRibbonControl)
ActiveChart.ChartArea.Copy
ActiveSheet.Range("A1").Select
ActiveSheet.Pictures.Paste.Select
End Sub

Sub AddChartTitle(control As IRibbonControl)
Dim i As Variant
i = InputBox("Please enter your chart title", "Chart Title")
On Error GoTo Last
ActiveChart.SetElement (msoElementChartTitleAboveChart)
ActiveChart.ChartTitle.Text = i
Last:
Exit Sub
End Sub


