Attribute VB_Name = "Module1"
Sub Yearly_Stock_Info()
'   Dim finish As Boolean
    Dim ticker_volume, grt_TVol, last_row As Long
    Dim last_column, index As Integer
    Dim oprice, clprice, percIncr, percDecr As Double
    Dim ticker, incrTicker, decrTicker, grt_TVolTicker As String
    percIncr = 0
    percDecr = 0
    index = 2
    ticker_volume = 0
'   MsgBox (Cells(Rows.Count, 1).End(xlUp).Row)
'  Loops through all worksheets and splits them in  a for loop
    For Each ws In Worksheets
        last_column = ws.Cells(1, Columns.Count).End(xlToLeft).Column + 1
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
       ' Set the columns and all the related starting data for first ticker
        ws.Cells(1, 1 + last_column).Value = "Ticker"
        ws.Cells(1, 2 + last_column).Value = "Yearly change"
        ws.Cells(1, 3 + last_column).Value = "Percent Change"
        ws.Cells(1, 4 + last_column).Value = "Total Stock Volume"
        ws.Cells(1, 8 + last_column).Value = "Ticker"
        ws.Cells(1, 9 + last_column).Value = "Value"
        ws.Cells(2, 7 + last_column).Value = "Greatest % Increase"
        ws.Cells(3, 7 + last_column).Value = "Greatest % Decrease"
        ws.Cells(4, 7 + last_column).Value = "Greatest Total Volume"
        oprice = ws.Cells(2, 3).Value
        ticker = ws.Cells(2, 1).Value
        ticker_volume = ticker_volume + ws.Cells(2, 7).Value
        For i = 3 To last_row
            If Not ticker = ws.Cells(i, 1) Then
                cprice = ws.Cells(i - 1, 6).Value
                ws.Cells(index, last_column + 1).Value = ticker
                ws.Cells(index, last_column + 2).Value = oprice - cprice
                If oprice - cprice >= 0 Then
                    ws.Cells(index, last_column + 2).Interior.ColorIndex = 10
                Else
                    ws.Cells(index, last_column + 2).Interior.ColorIndex = 30
                End If
                If oprice = 0 Then
                    ws.Cells(index, last_column + 3).Value = cprice
                Else
                    ws.Cells(index, last_column + 3).Value = (oprice - cprice) / oprice
                End If
                ws.Cells(index, last_column + 3).Style = "Percent"
                ws.Cells(index, last_column + 4).Value = ticker_volume
                ' Set the greatest volume, highest increase and highest decrease per pass through of sheet
                If ticker_volume > grt_TVol Then
                    grt_TVol = ticker_volume
                    grt_TVolTicker = ticker
                End If
                If percDecr > ws.Cells(index, last_column + 3).Value Then
                    percDecr = ws.Cells(index, last_column + 3).Value
                    decrTicker = ticker
                End If
                If ws.Cells(index, last_column + 3).Value > percIncr Then
                    percIncr = ws.Cells(index, last_column + 3).Value
                    incrTicker = ticker
                End If
                ' Reset new ticker initial values and start volume sum for new ticker
                ticker_volume = 0
                ticker_volume = ticker_volume + ws.Cells(i, 7)
                index = index + 1
                ticker = ws.Cells(i, 1).Value
                oprice = ws.Cells(i, 3).Value
            Else
                ' accumulate ticker volume
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
            End If
        Next i
        ' Gather the data for the last row and output the information for the last ticker
        cprice = ws.Cells(last_row, 6)
        ws.Cells(index, last_column + 1) = ticker
        ws.Cells(index, last_column + 2) = oprice - cprice
        If oprice = 0 Then
            ws.Cells(index, last_column + 3) = cprice
        Else
            ws.Cells(index, last_column + 3) = (oprice - cprice) / oprice
        End If
        ws.Cells(index, last_column + 3).Style = "Percent"
        ws.Cells(index, last_column + 4) = ticker_volume
        ' Reset for the next worksheet
        ticker_volume = 0
        index = 2
        ' Set the values for the greatest total and highest percent increase and decrease
        ws.Cells(2, 8 + last_column).Value = incrTicker
        ws.Cells(3, 8 + last_column).Value = decrTicker
        ws.Cells(4, 8 + last_column).Value = grt_TVolTicker
        ws.Cells(2, 9 + last_column).Value = percIncr
        ws.Cells(2, 9 + last_column).Style = "Percent"
        ws.Cells(3, 9 + last_column).Value = percDecr
        ws.Cells(3, 9 + last_column).Style = "Percent"
        ws.Cells(4, 9 + last_column).Value = grt_TVol
    Next
            
            
'    MsgBox (Worksheets(1).Cells(1, 1))
'    finish = Cells(1, 9).Value = ""
'    Do While finish
'
'    Exit Do
'    Worksheets(1).cells(i,j).Font
End Sub
' ws.cells(i,1)<>ws.cells(i+1,1) true if they are the same
' Create the needed variables
' Get the length of the rows
' Gather the start and end dates for the ticker
' Set the cells for each column to ticker, Yearly change, Percent Change and Total Stock Volume
' Calculate the stats for Yearly change, Percent Change and Total Stock Volume
' Gather the greatest % increase and decrease and total volume


