Attribute VB_Name = "Module1"
Sub Stock()

  Dim Ticker As String

  Dim Volume_Total As Double
  
  Dim Yearly_change As Double
  
  Dim Percent_change As Double
    
  Dim Open_price As Double
  
  Dim Close_price As Double
  
  Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

FirstTime = 0

Volume_Total = 0
    
  For Each ws In Worksheets
  
    'Header for summary table
      ws.Range("I1, O1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("P1").Value = "Value"
      ws.Range("N2").Value = "Greatest % Increase"
      ws.Range("N3").Value = "Greatest % Decrease"
      ws.Range("N4").Value = "Greatest Total Volume"

      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all stock transactions
    For i = 2 To LastRow

      If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
    
         FirstTime = FirstTime + 1

         Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        
    'Find the Open Price
          If FirstTime = 1 Then
          
              Open_price = ws.Cells(i, 3)
              
          Else
        
          End If
              
    'If the cell immediately following a row is not the same stock name...
     Else
 
         Ticker = ws.Cells(i, 1).Value

         Volume_Total = Volume_Total + ws.Cells(i, 7).Value

         Close_price = ws.Cells(i, 6).Value
      
            If Open_price <> 0 Then

               Yearly_change = Close_price - Open_price
      
               Percent_change = Yearly_change / Open_price

            Else

               Yearly_change = 0

               Percent_change = 0

            End If
      ' Print the stock name and other values in the Summary Table
          ws.Range("I" & Summary_Table_Row).Value = Ticker

          ws.Range("L" & Summary_Table_Row).Value = Volume_Total
      
          ws.Range("J" & Summary_Table_Row).Value = Yearly_change
       
          ws.Range("K" & Summary_Table_Row).Value = Percent_change
          
          ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
          
      ' Color formatting the yearly_change values
             If ws.Cells(Summary_Table_Row, 10).Value > 0 Then

                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4

             Else
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3

             End If

      ' Reset the volume_Total and FirstTime for next stock
        Volume_Total = 0

        FirstTime = 0
        
        Summary_Table_Row = Summary_Table_Row + 1

      End If

    Next i
    
    ' Reset the Summary_Table_Row for next sheet
    
    Summary_Table_Row = 2
    
    ' Set cell format as percent
    
    ws.Range("P2:P3").NumberFormat = "0.00%"
        
    ' Find Greatest % Increase
    
    Greatest_Increase = WorksheetFunction.Max(ws.Range("K:K"))
    
    ws.Range("P2").Value = Greatest_Increase
    
    ' Find Ticker name for the Greatest % Increase
    Ticker_Increase = WorksheetFunction.Match(ws.Range("P2"), ws.Range("K:K"), 0)
    
    ws.Range("O2").Value = ws.Range("I" & Ticker_Increase).Value
    
    
    ' Find Greatest % Decrease
    Greatest_Decrease = WorksheetFunction.Min(ws.Range("K:K"))
    
    ws.Range("P3").Value = Greatest_Decrease
    
    Ticker_Decrease = WorksheetFunction.Match(ws.Range("P3"), ws.Range("K:K"), 0)
    
    ws.Range("O3").Value = ws.Range("I" & Ticker_Decrease).Value
    
       
    ' Find Greatest Total Volume
   Greatest_Volume = WorksheetFunction.Max(ws.Range("L:L"))
    
   ws.Range("P4").Value = Greatest_Volume
   
   Ticker_Volume = WorksheetFunction.Match(ws.Range("P4"), ws.Range("L:L"), 0)
    
   ws.Range("O4").Value = ws.Range("I" & Ticker_Volume).Value

        
  Next ws
    
End Sub

