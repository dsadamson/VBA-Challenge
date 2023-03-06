Attribute VB_Name = "Module1"
Sub stock_data():

For Each ws In Worksheets

' declare variables
Dim WSname As String
Dim ticker As String
Dim x As Double
Dim y As Double
Dim vol As Double
Dim last_row As Long
Dim j As Long
Dim year_open As Double
Dim year_close As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim my_range As Range

' initialize variables
vol = 0
j = 0
WSname = ws.Name




'create summary headers
ws.Range("I1").Value = "ticker"
ws.Range("J1").Value = "yearly change"
ws.Range("K1").Value = "percentage change"
ws.Range("L1").Value = "total stock volume"
ws.Range("P1").Value = "ticker"
ws.Range("Q1").Value = "value"
ws.Range("O2").Value = "greatest % increase"
ws.Range("O3").Value = "greatest % decrease"
ws.Range("O4").Value = "greatest total volume"

' format summary tables
Columns("I:Q").Select
Columns("I:Q").EntireColumn.AutoFit

' find value for last row
last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

' find opening value
year_open = ws.Range("C2").Value



' loop through all stock transactions
For I = 2 To last_row

    'when it's the same
    
      If ws.Cells(I, 1).Value = ws.Cells(I + 1, 1).Value Then
        'add stock volume to vol
        vol = vol + ws.Cells(I, 7).Value
    
    Else

        'record ticker in summary table
        ws.Range("I" & 2 + j).Value = ws.Cells(I, 1).Value
        
    
        
        'Find value for year_close
        year_close = ws.Cells(I, 6).Value
        
        'Subtract year_open from year_close; place in Range("J" & 2+j)
   ws.Range("J" & 2 + j).Value = year_close - year_open
            
            'Add color fill to year change
            If ws.Range("J" & 2 + j).Value < 0 Then
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            Else
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            End If
    
        

        'Find percent change; place in Range("K" & 2+j)
    ws.Range("K" & 2 + j).Value = (year_close - year_open) / year_open
       
       'Find greatest % increase and decrease
            
       ws.Range("Q2") = Application.WorksheetFunction.Max(ws.Range("K2:K3001").Value)
            'find ticker for stock w/ greatest % increase
            ws.Range("P2").Value = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K3001").Value, 0)
            ' Note: '+1' included after range value because match function gives number of rows BEFORE the matching value
            ws.Range("P2").Value = ws.Cells(ws.Range("P2").Value + 1, 9).Value
          
       ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K2:K3001").Value)
            ws.Range("P3").Value = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K3001").Value, 0)
            
            ws.Range("P3").Value = ws.Cells(ws.Range("P3").Value + 1, 9).Value
            
        'format percent change value
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        
             
        'Reset year_open and find new value
    year_open = ws.Cells(I + 1, 3).Value
    
        
        'record volume in summary table
            'add stock volume to vol
        vol = vol + ws.Cells(I, 7).Value
        
        ws.Range("L" & 2 + j).Value = vol
        
        'Find greatest volume
        ws.Range("Q4") = Application.WorksheetFunction.Max(ws.Range("L2:L3001").Value)
            ws.Range("P4") = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L3001").Value, 0)
            ws.Range("P4").Value = ws.Cells(ws.Range("P4").Value + 1, 9).Value
        
        'reset vol
        vol = 0
        
        'reset close
        year_close = 0
        
        'increment j by 1
        j = j + 1
        
       

    End If

Next I

Next ws

End Sub


