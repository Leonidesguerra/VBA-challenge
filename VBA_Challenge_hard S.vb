Sub Stock_Summary()

    Dim WS As Worksheet
    For Each WS In Worksheets

        Dim tickrow As Long
        Dim tickrow2 As Long
        Dim I As Long
        dim j as long
        Dim ticker As Long
        Dim opvalue As Double
        Dim volume As Double 'volume is decleared as double because in some cases the total value excedes long capacity
        Dim gi As Double 'greatest increase
        Dim gd As Double  'greatest decrease
        Dim gv As Double  'greatest volume
        Dim tgi As String  'ticker for greatest increase
        Dim tgd As String  'ticker for greatest decrease
        Dim tgv As String  'ticker for greatest volume
    

        WS.Range("I1").Value = "Ticker"
        WS.Range("J1").Value = "Yearly Change"
        WS.Range("K1").Value = "% change"
        WS.Range("L1").Value = "Total Stock Volume"
        ticker = 2
        gi = 0
        gd = 0
        gv = 0
    
    
    
        'determine the number of rows the data table
        tickrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        volume = 0
    
        'The code inside this for will generate the summary table
        For I = 2 To tickrow
   
            'This if will determine if the row in wich it is working is the first for each specifick ticker e.g. A or AA-B etc.
            If WS.Cells(I, 1) <> WS.Cells(I - 1, 1) Then
                WS.Cells(ticker, 9).Value = WS.Cells(I, 1).Value
                opvalue = WS.Cells(I, 3).Value
                volume = WS.Cells(I, 7).Value
          
            
            'This elseif determines if this row is the last row for each particular ticker
            'If the row is the last row for a particular ticker the it calculates the yearly change and the % change
            ElseIf WS.Cells(I, 1) <> WS.Cells(I + 1, 1) Then
            WS.Cells(ticker, 10).Value = WS.Cells(I, 6).Value - opvalue
            
                'This if assigns different colors  for negative and positive values of the yearly change
                If WS.Cells(ticker, 10).Value < 0 Then
                WS.Cells(ticker, 10).Interior.ColorIndex = 3
                Else
                WS.Cells(ticker, 10).Interior.ColorIndex = 4
                End If
            
                'Chek if openvalue = 0 to prevent division by 0 in order to calculate the % change
                If opvalue = 0 Then
                opvalue = 1
                End If
                WS.Cells(ticker, 11).Value = Format(WS.Cells(ticker, 10) / opvalue, "percent")
                WS.Cells(ticker, 12) = volume + WS.Cells(I, 7).Value
                ticker = ticker + 1
                volume = 0
        
            'This else acumulates the volume value of every day per ticker
            Else
            volume = (WS.Cells(I, 7).Value) + volume
            
            End If
        Next I


        'BONUS:
        'tickrow2 gives the row count in the summary table
        tickrow2 = WS.Cells(Rows.Count, 9).End(xlUp).Row
   
        'this for returns the greatest values for: yearly decrease and increase as wella as total volume
        For j = 2 To tickrow2

            'if the value is positive it means it is an increase then compares it to the previous value stored in gi variable
            If WS.Cells(j, 11) > 0 And WS.Cells(j, 11) > gi Then
                gi = WS.Cells(j, 11).Value
                tgi = WS.Cells(j, 9).Value
            
            'if the value is negative it means it is a decrease then compares it to the previous value stored in gd variable
            ElseIf WS.Cells(j, 11) < 0 And WS.Cells(j, 11) < gd Then
                gd = WS.Cells(j, 11).Value
                tgd = Cells(j, 9).Value
            End If

            'this if gives in return the greatest value for total volume
            If WS.Cells(j, 12) > gv Then
                gv = WS.Cells(j, 12).Value
                tgv = WS.Cells(j, 9).Value
            End If

        Next j


        ' the folloing code prints, the greatest values obtained previously, in a list
        WS.Columns("O").ColumnWidth = 22
        WS.Columns("Q").ColumnWidth = 11
   
        WS.Range("P1") = "Ticker"
        WS.Range("Q1") = "Value"
   
        WS.Range("O2") = "Greatest % Increase"
        WS.Range("P2") = tgi
        WS.Range("Q2") = Format(gi, "percent")
   
        WS.Range("O3") = "Greatest % Decrease"
        WS.Range("P3") = tgd
        WS.Range("Q3") = Format(gd, "percent")
   
        WS.Range("O4") = "Greatest % Total Volume"
        WS.Range("P4") = tgv
        WS.Range("Q4") = gv
   
    Next WS
End Sub


