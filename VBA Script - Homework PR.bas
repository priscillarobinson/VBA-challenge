Attribute VB_Name = "Module1"
Sub Alph_testing()
'Define Variables
    'Loop through worksheets
    Dim WS As Worksheet
    
    'Row count
    Dim LastRow As Long
   
    'Loop through all worksheets
    For Each WS In Worksheets
        
        'Add Headers
        WS.Cells(1, "I").Value = "Ticker"
        WS.Cells(1, "J").Value = "Yearly Change"
        WS.Cells(1, "K").Value = "Percent Change"
        WS.Cells(1, "L").Value = "Total Stock Volume"


        'Define variables
      
        Dim Ticker_Name As String
        Dim ticker_amount As Integer
        Dim total_vol As Double
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_price_change As Double
        Dim percent_change As Double
        
        'Reset ticker amount
        ticker_amount = 0
             
        'Last row in column A
        LastRow = WS.Cells(Rows.Count, "A").End(xlUp).Row
            
        'Loop through each row of worksheets excluding first row
        For I = 2 To LastRow
                
            'Set Ticker name starting point
            Ticker_Name = WS.Cells(I, "A").Value
                    
            'Ticker opening price
            If open_price = 0 Then
                open_price = WS.Cells(I, "C").Value
            End If
                    
            'Ticker total volume
            total_vol = total_vol + WS.Cells(I, "G").Value
                    
                   
            'If different ticker name
            If WS.Cells(I + 1, 1).Value <> Ticker_Name Then
            
                    ticker_amount = ticker_amount + 1
                    WS.Cells(ticker_amount + 1, "I") = Ticker_Name
                    
                    'Ticker closing price
                    close_price = WS.Cells(I, "F").Value
                    
    
                    'yearly change
                    yearly_change = close_price - open_price
                        
                    'print yearly change
                    WS.Cells(ticker_amount + 1, "J").Value = yearly_change
                    
                    'format yearly change
                    If yearly_change >= 0 Then
                        WS.Cells(ticker_amount + 1, "J").Interior.ColorIndex = 4
                    
                    Else
                        WS.Cells(ticker_amount + 1, "J").Interior.ColorIndex = 3
                    
                    End If
                        
        
                    'percent change
                    If open_price = 0 Then
                        percent_change = 0
                    
                    Else
                       percent_change = (yearly_change / open_price)
                       
                    End If
                    
                    'print percent change as %
                    WS.Cells(ticker_amount + 1, "K").Value = Format(percent_change, "Percent")
                    
                    'print total stock volume
                    WS.Cells(ticker_amount + 1, "L").Value = total_vol

                'Reset values
                yearly_change = 0
                total_vol = 0
                open_price = 0
                Ticker_Name = ""
                close_price = 0
                percent_change = 0
   
             End If
    Next I
    
    'Define variables
    'Greatest increase
    Dim greatest_increase As String
    Dim GI_Value As Double
    
    'Greatest decrease
    Dim greatest_decrease As String
    Dim GD_Value As Double
    
    'Greatest volume
    Dim greatest_volume As String
    Dim GTV As Double
    
       
    'Headers for summary table
    WS.Cells(2, "N").Value = "Greatest % Increase"
    WS.Cells(3, "N").Value = "Greatest % Decrease"
    WS.Cells(4, "N").Value = "Greatest Total Volume"
    WS.Cells(1, "O").Value = "Ticker"
    WS.Cells(1, "P").Value = "Value"
    


  
    LastRow = WS.Cells(Rows.Count, "I").End(xlUp).Row
    
   'Reset Variables
    GI_Value = 0
    GD_Value = 0
    GTV = 0
   
    For j = 2 To LastRow

        
            'Greatest % increase
            If WS.Cells(j, "K").Value > GI_Value Then
                greatest_increase = Cells(j, "K").Value
                GI_Value = WS.Cells(j, "K").Value
                greatest_increase = WS.Cells(j, "I").Value
            End If
        
            'Greatest % decrease
            If WS.Cells(j, "K").Value < GD_Value Then
                greatest_decrease = WS.Cells(j, "K").Value
                GD_Value = WS.Cells(j, "K").Value
                greatest_decrease = WS.Cells(j, "I").Value
            End If
        
            'Greatest total volume
            If WS.Cells(j, "L").Value > GTV Then
                greatest_volume = WS.Cells(j, "L").Value
                GTV = WS.Cells(j, "L").Value
                greatest_volume = WS.Cells(j, "I").Value
            End If

        
    Next j
    
    'Print values
    WS.Cells(2, "P").Value = Format(GI_Value, "Percent")
    WS.Cells(2, "O").Value = greatest_increase
    WS.Cells(3, "P").Value = Format(GD_Value, "Percent")
    WS.Cells(3, "O").Value = greatest_decrease
    WS.Cells(4, "P").Value = Format(GTV, "General Number")
    WS.Cells(4, "O").Value = greatest_volume
    
    
WS.Range("N:P").HorizontalAlignment = xlCenter
WS.Range("N:P").Columns.AutoFit
    
Next WS

End Sub
