VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddGoalForm 
   Caption         =   "New Goal"
   ClientHeight    =   4070
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7490
   OleObjectBlob   =   "AddGoalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddGoalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubmitBtn_Click()

    Set WB = ThisWorkbook
    Set WS = WB.Worksheets("Budget&Goals")
    
    intRow = 2
    
    Do While (WS.Cells(intRow, "F") <> "")
        intRow = intRow + 1
    Loop
    
    'Check description
    If txtDescription.Value <> "" Then
        
        'Check Date
        If (txtDay.Value <> "" And txtMonth.Value <> "" And txtYear.Value <> "") Then
        
            'Check required savings
            If txtSavings.Value <> "" Then
            
                'Input achieve by date
                WS.Cells(intRow, "F").Value = txtYear.Value + "-" + txtMonth.Value + "-" + txtDay.Value
        
                'Format cell so Excel recognizes a date
                WS.Cells(intRow, "F").NumberFormat = "yyyy-mm-dd;@"
                
                'Input Description
                WS.Cells(intRow, "G").Value = txtDescription.Value
                
                'Input Required savigs
                WS.Cells(intRow, "H").Value = txtSavings.Value
                WS.Cells(intRow, "H").NumberFormat = "$#,##0.00"
            
            Else
                MsgBox ("Please enter a valid savings amount.")
            End If
        
        'Invalid Date
        Else
            MsgBox ("Please enter a valid date.")
        End If
    
    'Invalid Description
    Else
        MsgBox ("Please enter a valid description.")
    End If
    
End Sub
