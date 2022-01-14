VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddExpenseForm 
   Caption         =   "New Expense"
   ClientHeight    =   6300
   ClientLeft      =   -190
   ClientTop       =   -750
   ClientWidth     =   8930.001
   OleObjectBlob   =   "AddExpenseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddExpenseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    'Set up drop-down menu for form
    With AddExpenseForm.cboxCategory
        .AddItem "Food"
        .AddItem "Academic"
        .AddItem "Entertainment"
        .AddItem "Other"
    End With
    
    With AddExpenseForm.cboxFrequency
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
        
            'Test value of Category combobox
            If (cboxCategory.Value <> "") Then
    
                'Go through rows, if they contain data, increment
                Do While (WS.Cells(intRow, "B") <> "")
                
                    'Increment row counter
                    intRow = intRow + 1
                
                Loop
                                
                'Write date into cell
                WS.Cells(intRow, "B") = txtYear.Value + "-" + txtMonth.Value + "-" + txtDay.Value
        
                'Format cell so Excel recognizes a date
                WS.Cells(intRow, "B").NumberFormat = "yyyy-mm-dd;@"
                
                'Write item into cell
                WS.Cells(intRow, "C") = txtItem.Value
                
                'Write category into cell
                WS.Cells(intRow, "D") = cboxCategory.Value
                
                'Write description into cell
                WS.Cells(intRow, "E") = txtDescription.Value
                
                'Write priority into cell
                If lowBtn.Value = True Then WS.Cells(intRow, "F") = "Low"
                If medBtn.Value = True Then WS.Cells(intRow, "F") = "Medium"
                If highBtn.Value = True Then WS.Cells(intRow, "F") = "High"
                
                'Write Amount into cell
                WS.Cells(intRow, "G") = (txtAmount.Value) * -1
                WS.Cells(intRow, "G").NumberFormat = "$#,##0.00"
                
                'Write Frequency into cell
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
                WS.Cells(intRow, "H").Value = freq
                
                'Update yearly cost
                Worksheets("Expenses&Incomes").Range("I" & intRow).Formula = "=G" & intRow & "*H" & intRow
                WS.Cells(intRow, "I").NumberFormat = "$#,##0.00"
                
                
                
                'Write date into 'Expenses&Incomes-Expanded' Sheet
                
                Set WS = Worksheets("Expenses&Incomes - Expanded")
                Worksheets("Expenses&Incomes - Expanded").Activate
                
                intRow = 2
                Do While (WS.Cells(intRow, "B") <> "")
                    intRow = intRow + 1
                Loop
                
                WS.Cells(intRow, "B").Value = txtYear.Value + "-" + txtMonth.Value + "-" + txtDay.Value
                WS.Cells(intRow, "B").NumberFormat = "yyyy-mm-dd;@"
                
                WS.Cells(intRow, "C").Value = txtItem.Value
                
                WS.Cells(intRow, "D") = cboxCategory.Value
                
                WS.Cells(intRow, "E") = txtDescription.Value
                
                If lowBtn.Value = True Then WS.Cells(intRow, "F") = "Low"
                If medBtn.Value = True Then WS.Cells(intRow, "F") = "Medium"
                If highBtn.Value = True Then WS.Cells(intRow, "F") = "High"
                
                'Write Amount into cell
                WS.Cells(intRow, "G") = (txtAmount.Value)
                WS.Cells(intRow, "G").NumberFormat = "$#,##0.00"
                
                step = 365 / freq
                
                Do While (WS.Cells(intRow, "B").Value + step) < "2026-04-01"
                    intRow = intRow + 1
                    
                    WS.Cells(intRow, "B").Value = WS.Cells(intRow - 1, "B").Value + step
                    
                    WS.Cells(intRow, "C").Value = txtItem.Value
                    
                    WS.Cells(intRow, "D") = cboxCategory.Value
                    
                    WS.Cells(intRow, "E") = txtDescription.Value
                    
                    If lowBtn.Value = True Then WS.Cells(intRow, "F") = "Low"
                    If medBtn.Value = True Then WS.Cells(intRow, "F") = "Medium"
                    If highBtn.Value = True Then WS.Cells(intRow, "F") = "High"
                    
                    'Write Amount into cell
                    WS.Cells(intRow, "G") = (txtAmount.Value)
                    WS.Cells(intRow, "G").NumberFormat = "$#,##0.00"
                    
                Loop
                
                ClearOutput
                
                Worksheets("Expenses&Incomes").Activate
                
            Else
                'Give error for no category
                MsgBox ("Please select a category")
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
