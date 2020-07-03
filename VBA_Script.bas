Attribute VB_Name = "Module1"
Sub ticker_summaries()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS AND ADD HEADERS
    ' --------------------------------------------
    For Each ws In Worksheets
        ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        ' MsgBox WorksheetName
        
    
     ' Add the word Ticker to the First Column Header
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
            ' --------------------------------------------
            ' LOOP THROUGH ALL TICKERS
            ' --------------------------------------------
            'set variable for holding ticker
            Dim ticker As String
            'set variable for holding yearly change
            Dim opening_price As Double
            Dim ticker_rownum As Double
            Dim yearly_change As Double
            'set variable for holding percent change
            Dim percent_change As Single
            
                  
            'set variable for holding total stock volume
            Dim total_s_volume As Double
            'keep track of each ticker location
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
            
            'loop through all tickers
     
             For i = 2 To lastRow
            
                ' Check if we are still within the same ticker, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                  ' Set the ticker name
                  ticker = ws.Cells(i, 1).Value
                  'set the opening price & closing price
            
                  ' Add to the ticker Total
                  total_s_volume = total_s_volume + ws.Cells(i, 7).Value
                  'calculate yrly change this doesn't seem right but I will hold it for now
                  yearly_change = (ws.Cells(i, 6).Value - opening_price)
                  'calculate % change
                  If opening_price <> 0 Then
                        percent_change = yearly_change / opening_price
                        Else
                        percent_change = 0
                  End If
                  ' Print the ticker in the Summary Table
                  ws.Range("I" & Summary_Table_Row).Value = ticker
                  'print the yearly change in Summary table
                  ws.Range("J" & Summary_Table_Row).Value = yearly_change
                  'print percent change in summary table
                  ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percent_change, 2)
                  ' Print the ticker total to the Summary Table
                  ws.Range("L" & Summary_Table_Row).Value = total_s_volume
                  'contitional formatting for cell fill
                  If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
                    ' Set the Cell Colors to Green
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    Else
                    ' Set the Cell Colors to Red
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                   End If
    
                  ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
                  
                  ' Reset the ticker total
                  total_s_volume = 0
                  yearly_change = 0
                  ticker_rownum = 0
                  opening_price = 0
                  percent_change = 0
     
            
                ' If the cell immediately following a row is the same brand...
                Else
            
                  ' Add to the ticker total
                  total_s_volume = total_s_volume + ws.Cells(i, 7).Value
                  ticker_rownum = ticker_rownum + 1
                    If ticker_rownum = 1 Then
                        opening_price = ws.Cells(i, 3).Value
                    End If

                End If
        
            Next i
    'Challenge
        'add lables to column O starting at row 2
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
        'add labels to row 1
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
        'set variables to hold ticker, values
            Dim c_ticker As String
            Dim max_increase As Double
            Dim max_volume As Double
            Dim Max_decrease As Double
        'set values to variables
            c_lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            max_increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
            max_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
            Max_decrease = Application.WorksheetFunction.Min(ws.Range("K:k"))
        'set variable values to new cells
            ws.Cells(2, 17).Value = FormatPercent(max_increase, 2)
            ws.Cells(3, 17).Value = FormatPercent(Max_decrease, 2)
            ws.Cells(4, 17).Value = max_volume
        'loop through all c_tickers
            For j = 2 To c_lastRow
            
            
                ' Check if we are still within the same ticker, if it is not...
                If ws.Cells(j, 11).Value = max_increase Then
                 ' Set the ticker name
                  c_ticker = ws.Cells(j, 9).Value
                  ws.Cells(2, 16).Value = c_ticker
                End If
                If ws.Cells(j, 11).Value = Max_decrease Then
                 ' Set the ticker name
                  c_ticker = ws.Cells(j, 9).Value
                  ws.Cells(3, 16).Value = c_ticker
                End If
                If ws.Cells(j, 12).Value = max_volume Then
                 ' Set the ticker name
                  c_ticker = ws.Cells(j, 9).Value
                  ws.Cells(4, 16).Value = c_ticker
                End If
                
            Next j

    Next ws
    
    
End Sub


