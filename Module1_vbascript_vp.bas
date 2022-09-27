Attribute VB_Name = "Module1"
Sub tickerStats()

'define everything
Dim ws As Worksheet
Dim ticker As String
Dim volume As Double
    volume = 0
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets

    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup integer for loop
    Summary_Table_Row = 2
    

    'loop
    For I = 2 To ws.UsedRange.Rows.Count
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
            'find all the ticker values
            ticker = ws.Cells(I, 1).Value
                
            'find the volume data
            volume = volume + ws.Cells(I, 7).Value
            
            'Set ticker and volume
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 12).Value = volume


            'find year_open and year_close
            year_open = ws.Cells(I, 3).Value
            year_close = ws.Cells(I, 6).Value
            
            'Calculate Yearly Change and Percent Change
            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / year_close
            
            'Set year_open and year_close
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            
                'Conditional formating
                If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
                
                    'Set background color to red
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                
                Else
                
                    'Set background color to green
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                
                End If
                    
                'Calculate and write percent change in column K (#11)
                If ws.Cells(Summary_Table_Row, 3).Value <> 0 Then
                    percent_change = ((ws.Cells(Summary_Table_Row, 6).Value - ws.Cells(Summary_Table_Row, 3).Value) / ws.Cells(Summary_Table_Row, 3).Value)
                    
                    'Percent formating
                    ws.Cells(Summary_Table_Row, 11).Value = Format(percent_change, "Percent")
                    
                Else
                    
                    'percent formatting for 0
                    'ws.Cells(Summary_Table_Row, 11).Value = Format(0, "Percent")
                    
                End If

                'Add one to summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset volume
                volume = 0
                
              Else
                
                'add to volume total
                volume = volume + ws.Cells(I, 7).Value
        
        End If

    'finish loop
    Next I

'move to next worksheet
Next ws


End Sub
