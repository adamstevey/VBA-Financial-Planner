VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddSecurityForm 
   Caption         =   "Add Security Form"
   ClientHeight    =   2840
   ClientLeft      =   70
   ClientTop       =   300
   ClientWidth     =   5500
   OleObjectBlob   =   "AddSecurityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddSecurityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SubmitSecurity_Click()
    
    Set WB = ThisWorkbook
    Set WS = WB.Worksheets("Investments")
    
    intRow = 3
    
    If (tickerBox.Value <> "") Then
        
        If (sharesBox.Value <> "") Then
        
            Do While (WS.Cells(intRow, "C") <> "")
                
                intRow = intRow + 1
            
            Loop
            
            ' Input ticker symbol
            WS.Cells(intRow, "B") = tickerBox.Value
            
            Range("B" & intRow).ConvertToLinkedDataType ServiceID:=268435456, LanguageCulture:="en-US"
            WS.Cells(intRow, "B").NumberFormat = "$#,##0.00"
            
            ' Input # of shares
            WS.Cells(intRow, "D") = sharesBox.Value
            
            ' Update price column
            WS.Cells(intRow, "C").Formula = "=$B$" & intRow & ".Price"
            
            'Update Total Value Column
            WS.Cells(intRow, "E").Formula = "=C" & intRow & "*D" & intRow
            WS.Cells(intRow, "E").NumberFormat = "$#,##0.00"
        
        Else
        MsgBox ("Please enter a value for '#shares'")
        
        End If
    
    Else
        MsgBox ("Please enter a value for 'Ticker'")
    
    End If
        
        
    
End Sub
