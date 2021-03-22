Attribute VB_Name = "Module1"
Sub StockAnalysis_MultipleYear()

Dim ws As Worksheet
For Each ws In Worksheets
          
          ws.Activate
          Dim Ticker As String
        
          Dim Total_Stock_Volume As Double
          Total_Stock_Volume = 0
          greatest_increase = 0
          greatest_decrease = 0
          greatest_Volume = 0
            
          'Row line for summary table
          Dim Summary_Table_Row As Integer
          Summary_Table_Row = 2
          
          'Printing titles
          ws.Range("I1").Value = "Ticker"
          ws.Range("J1").Value = "Yearly Change"
          ws.Range("K1").Value = "Percentage Change"
          ws.Range("L1").Value = "Total Stock Volume"
          ws.Range("P1").Value = "Ticker"
          ws.Range("Q1").Value = "Value"
          ws.Range("O2").Value = "Greatest % Increase"
          ws.Range("O3").Value = "Greatest % decrease"
          ws.Range("O4").Value = "Greatest total volume"
          

          LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          Initial_Stock_Value = Cells(2, 3).Value
          
          For i = 2 To LastRow
            
            'Condition that trigger the summary calculations when we found a different ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
              'Obtatin ticker and it close value
              Ticker = Cells(i, 1).Value
              Close_Stock_Value = Cells(i, 6).Value

              
              'Calucale the yearly change
              Change_Stock_Value = Close_Stock_Value - Initial_Stock_Value
            
              'Calucale the yearly percentage change
              If Initial_Stock_Value = 0 Then
                Initial_Stock_Value = 1
                percentage_change_stock_value = (Close_Stock_Value - Initial_Stock_Value) / Initial_Stock_Value
                percentage_change_stock_value = 0
              Else
                percentage_change_stock_value = (Close_Stock_Value - Initial_Stock_Value) / Initial_Stock_Value
              End If
              
              
              'BONUS POINTS Calculating the greatest/lowest percentage change
              If percentage_change_stock_value > 0 Then
                If percentage_change_stock_value > greatest_increase Then
                greatest_increase = percentage_change_stock_value
                Range("P2").Value = Ticker
                Range("Q2").Value = greatest_increase
                End If
              Else
                If percentage_change_stock_value < greatest_decrease Then
                greatest_decrease = percentage_change_stock_value
                Range("P3").Value = Ticker
                Range("Q3").Value = greatest_decrease
                End If
              End If
                            
                           
              'Calculate total stock volume
              Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
              If Total_Stock_Volume > greatest_Volume Then
              greatest_Volume = Total_Stock_Volume
              Range("P4").Value = Ticker
              Range("Q4").Value = greatest_Volume
              End If
            
            
            
              ' Print Ticker in the Summary Table
              ws.Range("I" & Summary_Table_Row).Value = Ticker
              ws.Range("J" & Summary_Table_Row).Value = Change_Stock_Value
              ws.Range("K" & Summary_Table_Row).Value = percentage_change_stock_value
              ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume


              Summary_Table_Row = Summary_Table_Row + 1
              Initial_Stock_Value = Cells(i + 1, 3).Value
              Total_Stock_Volume = 0
              Close_Stock_Value = 0
              Change_Stock_Value = 0
              percentage_change_stock_value = 0
            Else
        
              ' Add to the Total Stock Volume
              Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
            End If
              
          Next i

Range("J:J").NumberFormat = "0.00"
Range("K:K").NumberFormat = "0.00%"
Range("Q2:Q3").NumberFormat = "0.00%"


Next ws

Worksheets(1).Activate
End Sub

