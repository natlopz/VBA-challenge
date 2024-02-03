Attribute VB_Name = "DataSet1"
'Multiple Year Stock Data

Sub StockData()

        'Variables for Data
    Dim ticker As String
    Dim startRow As Long
    Dim endRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalSVolume As LongLong
    Dim yearChange As Double
    Dim percentChange As Double
    Dim outputData As String
    
        'Setting Values for Data Variables
    totalSVolume = G
    
        'Variables for Greatest Increase, Decrease and Volume Value
    Dim gIncrease As Double
    Dim gDecrease As Double
    Dim gVolume As LongLong
    Dim tickergIncrease As String
    Dim tickergDecrease As String
    
    Dim tickergVolume As String
    

        'Setting Values for Greatest I, D, and V Variable
    gIncrease = 0
    gDecrease = 0
    gVolume = 0
    
        'Headers for Output Data
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    
        'Output Row
    Dim OutputRow As Long
    OutputRow = 2
    
        'Total Volume
    totalSVolume = 0
    
        'Looping through tickers
    For i = 2 To 735001
    If i = 2 Or Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
    If i > 2 Then
    closePrice = Cells(i - 1, 6).Value
    
        'Calculate Percent Change
    yearChange = closePrice - openPrice
    If openPrice <> 0 Then
    percentChange = (yearChange / openPrice) * 100
    Else
    percentChange = 0
    End If
    
        'Calculate Greatest Increase, Decrease and Volume
    
    If percentChange > gIncrease Then
    gIncrease = percentChange
    tickergIncrease = ticker
    End If
    
    If percentChange < gDecrease Then
    gDecrease = percentChange
    tickergDecrease = ticker
    End If
    
    If totalSVolume > gVolume Then
    gVolume = totalSVolume
    tickergVolume = ticker
    End If
    

        'Output the results into cells
    Cells(OutputRow, 9).Value = ticker
    Cells(OutputRow, 10).Value = yearChange
    Cells(OutputRow, 11).Value = Format(percentChange, "0.00") & "%"
    Cells(OutputRow, 12).Value = totalSVolume
    
    OutputRow = OutputRow + 1
    totalSVolume = 0
    End If
    
        'Handle the New Ticker
    ticker = Cells(i, 1).Value
    openPrice = Cells(i, 3).Value
    End If
    
        ' Accumulate volume for the current ticker
    totalSVolume = totalSVolume + Cells(i, 7).Value
        
    Next i
    
        ' Handle the Last Ticker
    ticker = Cells(i - 1, 1).Value
    closePrice = Cells(i - 1, 6).Value
    yearChange = closePrice - openPrice
    If openPrice <> 0 Then
        percentChange = (yearChange / openPrice) * 100
    Else
        percentChange = 0
    End If
    Cells(OutputRow, 9).Value = ticker
    Cells(OutputRow, 10).Value = yearChange
    Cells(OutputRow, 11).Value = Format(percentChange, "0.00") & "%"
    Cells(OutputRow, 12).Value = totalSVolume
    
        ' Formatting for Colors for Yearly Change and Percent Change
      Dim ychange As Range
      Dim pchange As Range
      
      Set ychange = Range("J2:J" & OutputRow)
      Set pchange = Range("K2:K" & OutputRow)
      
      For Each cell In ychange
      If cell.Value < 0 Then
        cell.Interior.Color = RGB(255, 0, 0)
        ElseIf cell.Value >= 0 Then
        cell.Interior.Color = RGB(0, 255, 0)
        
        End If
        
        Next cell
    
        For Each cell In pchange
        If cell.Value < 0 Then
        cell.Interior.Color = RGB(255, 0, 0)
        ElseIf cell.Value >= 0 Then
        cell.Interior.Color = RGB(0, 255, 0)
        
        End If
        
        Next cell
    
        'Headers for Greatest Values
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
        'Output for Greatest Values
    Range("Q2").Value = Format(gIncrease, "0.00") & "%"
    Range("Q3").Value = Format(gDecrease, "0.00") & "%"
    Range("Q4").Value = gVolume
    Range("P2").Value = tickergIncrease
    Range("P3").Value = tickergDecrease
    Range("P4").Value = tickergVolume
    
    
    
End Sub

