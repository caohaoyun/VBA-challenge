Attribute VB_Name = "Module1"
Sub stock_anallysis()
'to run thecodein all worksheets of this workbook

Dim num_of_sheets As Integer
Dim ws As Worksheet

num_of_sheets = ActiveWorkbook.Worksheets.count

For Each ws In Worksheets

'Name new columns and rows for analysis
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"

    ws.Cells(2, 17).Value = "Greatest % Increase"
    ws.Cells(3, 17).Value = "Greatest % Decrease"
    ws.Cells(4, 17).Value = "Greatest Total Volume"
    ws.Cells(1, 18).Value = "Ticker"
    ws.Cells(1, 19).Value = "Value"

'Find all of the tickers from column A
    Dim ticker As String
    Dim totalvolume As LongLong


    Dim row As Integer
    Dim i As Long
    Dim lastrow As Long
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim count As LongLong


    lastrow = ws.Cells(Rows.count, "A").End(xlUp).row
    row = 2
    count = 0
    totalvolume = 0
    'start loop for calculating the value
        For i = 2 To lastrow
    
    'verify if the ticker change to a new one
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
       
        
        'find all unique ticker names and put them into column k
            ticker = ws.Cells(i, 1).Value
            ws.Range("K" & row).Value = ticker
        
        'calculate yearly change and put them into column L
            yearlychange = ws.Cells(i, 6).Value - ws.Cells(i - count, 3).Value
            ws.Range("L" & row).Value = yearlychange
        
        'calculate percentage change using yearly change and open price, then put them into column M
            percentchange = (ws.Cells(i, 6).Value - ws.Cells(i - count, 3).Value) / IIf(ws.Cells(i - count, 3).Value = 0, 1, ws.Cells(i - count, 3).Value)
            ws.Range("M" & row).Value = percentchange
            ws.Range("M" & row).NumberFormat = "0.00%"
        
        'put total stock volume into column N
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            ws.Range("N" & row).Value = totalvolume
        
        
            count = 0
            row = row + 1
            totalvolume = 0
        Else
        'calculate the value for the new ticker
        
            ticker = ws.Cells(i, 1).Value
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            count = count + 1
    
        
        End If
        Next i
    
'color the year change column according to their increase or decrease
    For i = 2 To lastrow
        If ws.Cells(i, 12).Value > 0 Then
            ws.Cells(i, 12).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 12).Value < 0 Then
            ws.Cells(i, 12).Interior.ColorIndex = 3
    
        Else
        
        End If
        Next i
    
    
'calculate greatest increase, decrease and total volume

    Dim max As Double
    Dim min As Double
    Dim volumemax As LongLong
    Dim tickermax As String
    Dim tickermin As String
    Dim tickervolumemax As String






    max = 0
    min = 0

    For i = 2 To lastrow
        For Each cell In ws.Cells(i, 13)
            If ws.Cells(i, 13).Value > max Then
                max = ws.Cells(i, 13).Value
                tickermax = ws.Cells(i, 11).Value
            ElseIf ws.Cells(i, 13).Value < min Then
                min = ws.Cells(i, 13).Value
                tickermin = ws.Cells(i, 11).Value
            Else
            End If
        Next
    Next

    ws.Cells(2, 18).Value = tickermax
    ws.Cells(2, 19).Value = max
    ws.Cells(3, 18).Value = tickermin
    ws.Cells(3, 19).Value = min
   
    For i = 2 To lastrow
        For Each cell In ws.Cells(i, 14)
            If ws.Cells(i, 14).Value > volumemax Then
                volumemax = ws.Cells(i, 14).Value
                tickervolumemax = ws.Cells(i, 11).Value
            End If
        Next
    Next

    ws.Cells(4, 18).Value = tickervolumemax
    ws.Cells(4, 19).Value = volumemax
    
Next ws
End Sub

