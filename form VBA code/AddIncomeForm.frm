VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddIncomeForm 
   Caption         =   "New Income"
   ClientHeight    =   5910
   ClientLeft      =   -20
   ClientTop       =   -150
   ClientWidth     =   9710.001
   OleObjectBlob   =   "AddIncomeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddIncomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    'Set up drop-down menu for form
    With AddIncomeForm.cboxFrequency
        .AddItem "One time"
        .AddItem "Monthly"
        .AddItem "Biweekly"
        .AddItem "Weekly"
    End With
End Sub

Private Sub SubmitBtn_Click()

    'Set workbook and sheet
    Set WB = ThisWorkbook
    Set WS = WB.Worksheets("Expenses&Incomes")

    'start on second row (headers are first row)
    intRow = 2

    'Test value of Item textbox
    If (txtItem.Value <> "") Then
    
        'Test value of date textboxes
        If (txtDay.Value <> "" And txtMonth.Value <> "" And txtYear.Value <> "") Then
        
            'Test value of Amount Text box
            If (txtAmount.Value <> "") Then
    
                'Go through rows, if they contain data, increment
                Do While (WS.Cells(intRow, "L") <> "")
                
                    'Increment row counter
                    intRow = intRow + 1
                
                Loop
                                
                'Write date into cell
                WS.Cells(intRow, "L") = txtYear.Value + "-" + txtMonth.Value + "-" + txtDay.Value
        
                'Format cell so Excel recognizes a date
                WS.Cells(intRow, "L").NumberFormat = "yyyy-mm-dd;@"
                
                'Write item into cell
                WS.Cells(intRow, "M") = txtItem.Value
                
                'Write description into cell
                WS.Cells(intRow, "N") = txtDescription.Value
                
                'Write amount into cell
                WS.Cells(intRow, "O") = txtAmount.Value
                WS.Cells(intRow, "O").NumberFormat = "$#,##0.00"
                
                'Write freuency into cell
                If cboxFrequency.Value = "Monthly" Then
                    freq = 12
                ElseIf cboxFrequency.Value = "Biweekly" Then
                    freq = 26
                ElseIf cboxFrequency.Value = "One time" Then
                    freq = 1
                ElseIf cboxFrequency.Value = "Weekly" Then
                    freq = 52
                Else
                    freq = cboxFrequency.Value
                End If
                
                WS.Cells(intRow, "P") = freq
                
                'Update yearly income
                Worksheets("Expenses&Incomes").Range("Q" & intRow).Formula = "=O" & intRow & "*P" & intRow
                WS.Cells(intRow, "Q").NumberFormat = "$#,##0.00"
                
                
                'Write date into 'Expenses&Incomes-Expanded' Sheet
                
                Set WS = Worksheets("Expenses&Incomes - Expanded")
                Worksheets("Expenses&Incomes - Expanded").Activate
                
                intRow = 2
                Do While (WS.Cells(intRow, "J") <> "")
                    intRow = intRow + 1
                Loop
                
                WS.Cells(intRow, "J").Value = txtYear.Value + "-" + txtMonth.Value + "-" + txtDay.Value
                WS.Cells(intRow, "J").NumberFormat = "yyyy-mm-dd;@"
                
                WS.Cells(intRow, "K").Value = txtItem.Value
                WS.Cells(intRow, "L").Value = txtDescription.Value
                'Write Amount into cell
                WS.Cells(intRow, "M") = (txtAmount.Value)
                WS.Cells(intRow, "M").NumberFormat = "$#,##0.00"
                
                step = 365 / freq
                
                Do While (WS.Cells(intRow, "J").Value + step) < "2026-04-01"
                    intRow = intRow + 1
                    
                    WS.Cells(intRow, "J").Value = WS.Cells(intRow - 1, "J").Value + step
                    
                    WS.Cells(intRow, "K").Value = txtItem.Value
                    
                    WS.Cells(intRow, "L") = txtDescription.Value
                    
                    'Write Amount into cell
                    WS.Cells(intRow, "M") = (txtAmount.Value)
                    WS.Cells(intRow, "M").NumberFormat = "$#,##0.00"
                    
                Loop
                
                ClearOutput
                
                Worksheets("Expenses&Incomes").Activate
                
            Else
                'Give error for no Amount
                MsgBox ("Please enter a valid amount")
            End If
        
        Else
            'Give error message for no date
            MsgBox ("Please enter a valid date")
        End If
        
    Else
        'Give error message for no item
        MsgBox ("Please enter an item")
    End If
    
End Sub
