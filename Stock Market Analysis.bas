Attribute VB_Name = "Module1"
Sub Stock_Analysis()

'Declaration of all variables

Dim ws As Worksheet
Dim Ticker As String
Dim Year_Opening_Value As Double
Dim Year_Closing_Value As Double
Dim Yearly_Change As Double
Dim Total_StockVol As Double
Dim Change_Percentage As Double
Dim Ticker_Row As Integer
Dim Greatest_Increase As Double

For Each ws In Worksheets

'Column headers as per instructions
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Initialisation of variables
Ticker_Row = 2
Previous_Ticker_Row = 1
Total_StockVol = 0
        
'Finding the last row of data in the worksheet
Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Looping through data
For i = 2 To Lastrow
            
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
       Ticker = ws.Cells(i, 1).Value
            
       Previous_Ticker_Row = Previous_Ticker_Row + 1
            
       Year_Opening_Value = ws.Cells(Previous_Ticker_Row, 3).Value
            
       Year_Closing_Value = ws.Cells(i, 6).Value
            

      For j = Previous_Ticker_Row To i
      
        Total_StockVol = Total_StockVol + ws.Cells(j, 7).Value
                
    Next j
            
       If Year_Opening_Value = 0 Then
       
          Change_Percentage = Year_Closing_Value
                
       Else
      
         Yearly_Change = Year_Closing_Value - Year_Opening_Value
         
         Change_Percentage = Yearly_Change / Year_Opening_Value
                
       End If
            
        'Assigning Values to locations for data in the summary table and formatting percentage of changes
        ws.Cells(Ticker_Row, 9).Value = Ticker
        ws.Cells(Ticker_Row, 10).Value = Yearly_Change
        ws.Cells(Ticker_Row, 11).Value = Change_Percentage
        ws.Cells(Ticker_Row, 11).NumberFormat = "0.00%"
        ws.Cells(Ticker_Row, 12).Value = Total_StockVol
                
        Ticker_Row = Ticker_Row + 1
                
       'Setting values back to zero
       Total_StockVol = 0
       Yearly_Change = 0
       Change_Percentage = 0
            
       'Changing previuos ticker number to the i
       Previous_Ticker_Row = i
        
    End If
    
Next i
    
    'Conditional formatting yearly change percentages in columns "J" in the summary table
    'Finding the last row in the summary table to create the loop
    
    Lastrow_j = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
For j = 2 To Lastrow_j
            
    If ws.Cells(j, 10) >= 0 Then
                
     ws.Cells(j, 10).Interior.ColorIndex = 4
                    
    Else
                
     ws.Cells(j, 10).Interior.ColorIndex = 3
                
    End If
                
Next j
    
    ' Looping through the summary table to create analysis summary table for greatest increase, decrease and greatest total stock volume
    ' Percentage changes are located in the column "K", total stock volumes are in column "L"
        
    'Column headers as per instructions
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
              
    lastrow_k = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
                For i = 2 To lastrow_k
            
                    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K:K")) Then
                        
                         'Assigning the ticker symbol and value from sumarry table to the designated cell in the analysis summary  table
                         ws.Range("P2") = ws.Cells(i, 9).Value
                         ws.Range("Q2") = ws.Cells(i, 11).Value
                         ws.Range("Q2").NumberFormat = "0.00%"
                        
                    ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K:K")) Then
                        
                         'Assigning the ticker symbol and value from sumarry table to the designated cell in the analysis summary  table
                         ws.Range("P3") = ws.Cells(i, 9).Value
                         ws.Range("Q3") = ws.Cells(i, 11).Value
                         ws.Range("Q3").NumberFormat = "0.00%"
                        
                    ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L:L")) Then
                        
                         'Assigning the ticker symbol and value from sumarry table to the designated cell in the analysis summary  table
                         ws.Range("P4") = ws.Cells(i, 9).Value
                         ws.Range("Q4") = ws.Cells(i, 12).Value
                
                    End If
        
              Next i
             
        Next ws

End Sub



