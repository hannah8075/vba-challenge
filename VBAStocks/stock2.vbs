Sub stock()
  
  'Loop through all sheets
  For Each ws In Worksheets

    'Sort sheet
    'With ActiveSheet.sort
      ' .SortFields.Add Key:=Range("A1"), Order:=xlAscending
      ' .SortFields.Add Key:=Range("B1"), Order:=xlAscending
      ' .SetRange Range("A1:G" & lRow)
      ' .Header = xlYes
      ' .Apply
    'End With

    'create summary table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    ' Set an initial variable for holding the ticker name
    Dim Ticker As String

    ' Set an initial variable for holding the volume total per ticker
    Dim VolTotal As Double
    VolTotal = 0

    ' Keep track of the location for ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Set last row variable for data
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set Open and Close Price
    Dim OpenPrice, ClosePrice, YearlyDiff, PercentChange As Double
    OpenPrice = ws.Range("C2").Value

    ' Loop through all stocks
    For i = 2 To LastRow

      ' Check if we are still within the same ticker, if it is not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ' Set the Ticker
        Ticker = ws.Cells(i, 1).Value

        ' Add to the volume total
        VolTotal = VolTotal + ws.Cells(i, 7).Value

        ' Print the ticker in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker

        ' Print the volume total to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = VolTotal
        
        'Set Last Price
        ClosePrice = ws.Cells(i, 6).Value
        
        'Calculate yearly diff
        YearlyDiff = ClosePrice - OpenPrice
        ws.Range("J" & Summary_Table_Row).Value = YearlyDiff

      'Calculate percent change if open price is not 0
          If OpenPrice <> 0 Then
              PercentChange = YearlyDiff / OpenPrice
          Else
              PercentChange = 0
          End If
      ' Put Percent change in sumary table
          ws.Range("K" & Summary_Table_Row).Value = PercentChange
      
      'Format cell color based on percent change
          If PercentChange < 0 Then
              ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
          Else
              ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
          End If
          
          'Format Percent change into percentage
          ws.Range("K" & Summary_Table_Row) = Format(PercentChange, "0.00%")
      
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Set OpenPrice for next ticker
        OpenPrice = ws.Cells(i + 1, 3).Value
        
        ' Reset the volume Total
        VolTotal = 0

      ' If the cell immediately following a row is the same ticker...
      Else

        ' Add to the volume Total
        VolTotal = VolTotal + ws.Cells(i, 7).Value


      End If

    Next i
    
    
    'Set last row variable for summary table
    Dim SummaryLastRow As Long
    SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'Set max and min variables
    'Dim MaxPercent, MinPercent, MaxVol As Long

    MaxPercent = WorksheetFunction.Max(ws.Range("K2:K" & SummaryLastRow))
    MinPercent = WorksheetFunction.Min(ws.Range("K2:K" & SummaryLastRow))
    MaxVol = WorksheetFunction.Max(ws.Range("L2:L" & SummaryLastRow))

      For k = 2 To SummaryLastRow
          If ws.Range("K" & k) = MaxPercent Then
              ws.Range("P2").Value = ws.Range("I" & k).Value
              ws.Range("Q2").Value = MaxPercent
          ElseIf Range("K" & k) = MinPercent Then
              ws.Range("P3").Value = ws.Range("I" & k).Value
              ws.Range("Q3").Value = MinPercent
          ElseIf ws.Range("L" & k) = MaxVol Then
              ws.Range("P4").Value = ws.Range("I" & k).Value
              ws.Range("Q4").Value = MaxVol
          End If
      Next k
           
      'Format Percent change into percentage
      ws.Range("Q2") = Format(MaxPercent, "0.00%")
      ws.Range("Q3") = Format(MinPercent, "0.00%")
Next ws

End Sub

