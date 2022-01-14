Attribute VB_Name = "Module1"
Sub ClearExpenses()

    'Select the range
    Range("B3:I3").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents
    
    Worksheets("Expenses&Incomes - Expanded").Activate
    Range("B3:G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ClearOutput
    Worksheets("Expenses&Incomes").Activate
    
End Sub
Sub ClearIncomes()

    'Select the range
    Range("L3:Q3").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents
    
    Worksheets("Expenses&Incomes - Expanded").Activate
    Range("J3:M3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ClearOutput
    Worksheets("Expenses&Incomes").Activate
    
End Sub
Sub ClearInvestments()

    'Select the range
    Range("B3:E3").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents
    
End Sub

Sub ClearSavings()

    'Select the range
    Range("H3:K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents
    
End Sub

Sub ClearOutput()
    Worksheets("Output - Expenses&Incomes").Activate
    Range("D3:I3").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.ClearContents
    
    Range("K3:N3").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.ClearContents
    
    Range("A3").ClearContents
    Range("A5").ClearContents
    
    Application.DisplayAlerts = False
    
    For Each wksheet In ThisWorkbook.Worksheets
    If wksheet.Name = "Expenses&Incomes Charts" Then
        wksheet.Delete
    End If
    
    Next
    
End Sub

Sub ClearGoals()
    Range("F3:H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Sub

Sub ShowOutputForm()
    'Open the output userForm
    OutputForm.Show
End Sub
Sub ShowExpenseForm()
    'Open the 'Add Expense' userForm
    AddExpenseForm.Show
End Sub
Sub ShowIncomeForm()
    'Show the 'Add income' userform
    AddIncomeForm.Show
End Sub
Sub ShowSecurityForm()
    ' Open the New Security form
    AddSecurityForm.Show
End Sub

Sub ShowSavingsForm()
    'Show the 'New Savings Account' Form
    AddSavingsForm.Show
End Sub

Sub ShowGoalsForm()
    'Show the 'New Goal' Form
    AddGoalForm.Show
End Sub

Sub ShowSelectGoalForm()
    'Show the 'Select Goal' Form
    SelectGoalForm.Show
End Sub

Sub UpdateSavings()

    Set WB = ThisWorkbook
    Set WS = WB.Worksheets("Investments")
    
    intRow = 3
    
    Dim rate As Double
    Dim dif As Double
    Dim years As Integer
    Dim i As Integer
    
    Do While (WS.Cells(intRow, "J") <> "")
    
        'Find difference in time between today's date and last update
        dif = Date - WS.Cells(intRow, "J").Value
        
        'Get rate value for row
        rate = (WS.Cells(intRow, "I").Value)
        
        'Find number of years since last update
        years = dif \ 365
        
        For i = 1 To years
        
            'Update account value based on years since last update
            WS.Cells(intRow, "K") = WS.Cells(intRow, "K").Value + ((WS.Cells(intRow, "K").Value) * (rate))
    
        Next
        
        'Update account value based on days since last update
        WS.Cells(intRow, "K") = WS.Cells(intRow, "K").Value + ((WS.Cells(intRow, "K")) * ((dif Mod 365) * (rate / 365)))
        
        'Update 'As of date''
        WS.Cells(intRow, "J") = Date
        
        intRow = intRow + 1
        
    Loop
    
End Sub

Sub IncomeVsExpensesChart()
    
    Dim cht As ChartObject
    
    Set cht = Sheets("Expenses&Incomes").ChartObjects.Add(Left:=1330, Width:=300, Top:=145, Height:=250)
    cht.Chart.SetSourceData Source:=Sheets("Expenses&Incomes").Range("S1:T4")
    
    Worksheets("Expenses&Incomes").ChartObjects(1).Activate
    
    With ActiveChart
        .SetElement msoElementLegendNone
        
        .HasTitle = True
        .ChartTitle.Text = "Yearly Net Income"
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Caption = "Cashflow ($)"
        End With
        
    End With
    
    Dim s As Series
    Set s = ActiveChart.SeriesCollection("Value ($)")

    For i = 1 To s.Points.Count
        If s.Values(i) > 0 Then s.Points(i).Interior.Color = RGB(0, 255, 0)
        If s.Values(i) < 0 Then s.Points(i).Interior.Color = RGB(255, 0, 0)
    Next

End Sub

Sub CreatePivotTable()
'PURPOSE: Creates a brand new Pivot table on a new worksheet from data in the ActiveSheet
'Source: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim lrow As Integer
Dim x As Worksheet

'Determine the data range you want to pivot
  lrow = Worksheets("Output - Expenses&Incomes").Cells(Rows.Count, "D").End(xlUp).Row
  SrcData = Sheets("Output - Expenses&Incomes").Range("D2:I" & lrow).Address(ReferenceStyle:=xlR1C1)

'Create a new worksheet
  For Each wksheet In ThisWorkbook.Worksheets
    If wksheet.Name = "Expenses Charts" Then
        wksheet.Delete
    End If
  Next
  Set sht = Sheets.Add

'Where do you want Pivot Table to start?
  StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")

  pvt.PivotFields("Category").Orientation = xlRowField
  
  With ActiveSheet.PivotTables("PivotTable1").PivotFields("Amount ($)")
  .Orientation = xlDataField
  .Function = xlSum
  End With
  
  With pvt.PivotFields("Priority")
    .Orientation = xlColumnField
    .PivotItems("High").Position = 3
  End With
  
  ActiveSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleDark27"
  
  sht.Name = "Expenses&Incomes Charts"
  
  CreateGraph

End Sub

Sub CreateGraph()
    Dim cht As ChartObject
    
    Set cht = Sheets("Expenses&Incomes Charts").ChartObjects.Add(Left:=0, Width:=475, Top:=110, Height:=300)
    cht.Chart.SetSourceData Source:=Sheets("Expenses&Incomes Charts").Range("A1:E7")
    
    Worksheets("Expenses&Incomes Charts").ChartObjects(1).Activate
    
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Text = "Expenses by Category and Priority"
        .ClearToMatchStyle
        .ChartStyle = 213
        
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Caption = "Category"
        End With
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Caption = "Sum ($)"
        End With
        
    End With
    
    'Set Colors
    ActiveChart.SeriesCollection("Low").Select
    Selection.Interior.Color = RGB(0, 255, 0)
    ActiveChart.SeriesCollection("Medium").Select
    Selection.Interior.Color = RGB(255, 255, 0)
    ActiveChart.SeriesCollection("High").Select
    Selection.Interior.Color = RGB(255, 0, 0)

End Sub

Sub NetIncomeChart()
    
    Dim cht As ChartObject
    
    Set cht = Sheets("Expenses&Incomes Charts").ChartObjects.Add(Left:=550, Width:=475, Top:=75, Height:=300)
    cht.Chart.SetSourceData Source:=Sheets("Expenses&Incomes Charts").Range("K2:L4")
    
    Worksheets("Expenses&Incomes Charts").ChartObjects(2).Activate
    
    With ActiveChart
        .SetElement msoElementLegendNone
        .HasTitle = True
        .ChartTitle.Text = "Net Income (For Selected Period)"
        
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Caption = "($)"
        End With
        
    End With
    
    Dim s As Series
    Set s = ActiveChart.SeriesCollection("Series1")

    For i = 1 To s.Points.Count
        If s.Values(i) > 0 Then s.Points(i).Interior.Color = RGB(0, 255, 0)
        If s.Values(i) < 0 Then s.Points(i).Interior.Color = RGB(255, 0, 0)
    Next
    
   

End Sub

Sub SetChartsSheet()
'
' SetChartsSheet Macro
'

'
    Columns("K:K").ColumnWidth = 15.36
    Columns("L:L").ColumnWidth = 15.64
    Range("K1").Select
    ActiveCell.Formula2R1C1 = "Datatype"
    Range("L1").Select
    ActiveCell.Formula2R1C1 = "Total"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "Total expenses"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "Total incomes"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "Net Income:"
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "Total Incomes"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "Total Expenses"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=SUM('Output - Expenses&Incomes'!C[-3])"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=-SUM('Output - Expenses&Incomes'!C[-3])"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "=SUM('Output - Expenses&Incomes'!C[2])"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-2]C, R[-1]C)"
    Range("L2:L4").Select
    Range("L4").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("K2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("K3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("K4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399945066682943
        .PatternTintAndShade = 0
    End With
    Range("K2:K4").Select
    Range("K4").Activate
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("N8").Select
End Sub

