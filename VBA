Sub allsheets()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ' Activate current sheet
        ws.Activate
        ' Select cell A1
        Range("A1").Select
        ' calling subroutine stockdata()
        stockdata ws
    Next ws

End Sub
    
Sub stockdata(ws As Worksheet)

    Dim ticker As String
    Dim rng As Range
    Dim lastrow As Long
    Dim year As Integer
    Dim i As Long
    Dim list As Integer
    Dim start As Long
    Dim change As Double
    Dim percentchange As Double
    Dim totalstock As Double
    Dim maxincrease As Double
    Dim maxdecrease As Double
    Dim greatesttotalvolume As Double
    Dim numericvalue As Double
    Dim maxincreaseticker As String
    Dim maxdecreaseticker As String
    Dim greatestvolumeticker As String


    list = 2
    start = 2
    maxincrease = 0
    maxdecrease = 0
    greatesttotalvolume = 0

    
        'go through all values in column A
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        start = 2
        
    'i=2 since the first row are headers
    For i = 2 To lastrow
    


            'go through values until a different one is found, and print results
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            ' calculates the change between last closing price and 1st opening price
            change = (ws.Cells(i, 6).Value - ws.Cells(start, 3).Value)
            
            percentchange = change / ws.Cells(start, 3).Value
            
            ' prints the change in the according ticker row to the right
            
            If change > 0 Then
            
            ws.Cells(list, 10).Interior.ColorIndex = 4
            

            Else
            
            ws.Cells(list, 10).Interior.ColorIndex = 3
            
            End If
            
            
            ws.Cells(list, 10).Value = change
            
            ' prints the percentchange

            ws.Cells(list, 11).Value = percentchange
            ws.Cells(list, 11).NumberFormat = "0.00%"
            
              'last ticker value (before new one) is obtained
            ticker = ws.Cells(i, 1).Value
            
            ' prints the tickers list
            ws.Cells(list, 9).Value = ticker
            
            ' reset totalstock for each new ticker
            totalstock = 0
            
                'construct the range for the current ticker
            
                Set rng = ws.Range("G" & start & ":G" & i)
                
                'Sum up the volume for the current ticker

                For Each cell In rng
            
                    totalstock = totalstock + cell.Value
            
                Next cell
            
            ws.Cells(list, 12).Value = totalstock

            'prints tickers in the respective column (I)
   
            list = list + 1
             
            ' a +1 is added in order to jump to the next ticker's first row
             start = i + 1
                 

            End If
            
        
    Next i
    
        'Autofit columns width
        ws.Range("I1").Value = "Ticker"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        Set rng = ws.Range("I:Q")
        rng.Columns.AutoFit
        
        
        ' Greatest % increase and decrease will be searched for
        Set rng = ws.Range("k2:k" & lastrow)
                
            For Each cell In rng
            
            numericvalue = cell.Value
            
            
                 If numericvalue > maxincrease Then
                    
                    maxincrease = numericvalue

                    maxincreaseticker = ws.Cells(cell.Row, "I").Value
                    
                  End If
                  
                  
                    
                  If numericvalue < maxdecrease Then
                    
                    maxdecrease = numericvalue
                    
                    maxdecreaseticker = ws.Cells(cell.Row, "I").Value
                    
                   End If
                    
            Next cell
            
           ' prints the ticker corresponding to the greatest increase and decrease in %
           ws.Cells(3, 16).Value = maxdecreaseticker

           ws.Cells(2, 16).Value = maxincreaseticker
           
            
          ' prints the greatest % of increase and decrease, while formatting the cell to percentage
          ws.Cells(2, 17).Value = maxincrease
          ws.Cells(2, 17).NumberFormat = "0.00%"

          
          ws.Cells(3, 17).Value = maxdecrease
          ws.Cells(3, 17).NumberFormat = "0.00%"

             
        
        
                ' Greatest Total Volume will be searched for in a different column (L)
                Set rng = ws.Range("L2:L" & lastrow)
                
                
            For Each cell In rng
            
            numericvalue = cell.Value
            
            
                 If numericvalue > greatesttotalvolume Then
                    
                    greatesttotalvolume = numericvalue
                    greatestvolumeticker = ws.Cells(cell.Row, "I").Value
                    
                 
                    End If
                        
                 
            Next cell
            
          ' prints the greatest % of increase and decrease, while formatting the cell to percentage
          ws.Cells(4, 17).Value = greatesttotalvolume
          ws.Cells(4, 16).Value = greatestvolumeticker

     
        
        
End Sub


            
