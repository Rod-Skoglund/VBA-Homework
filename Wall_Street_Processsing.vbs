Sub ProcessTickerSymbols()

  '*******************************************************************************
  'Define variables
  '*******************************************************************************

  '*******************************************************************************
  'Current Ticker variables
  '*******************************************************************************
  Dim curTicker As String
  Dim curTickerTotalVolume As Double
  Dim curTickerOpen As Double
  Dim curTickerchange As Double
  Dim curTickerPercentChange As Double

  '*******************************************************************************
  'Track Greatest Increase, Decrease and Volume and the associated Ticker Symbol
  '*******************************************************************************
  Dim maxPercentIncreaseTicker As String
  Dim maxPercentIncrease As Double
  Dim maxPercentDecreaseTicker As String
  Dim maxPercentDecrease As Double
  Dim maxTotalVolumeTicker As String
  Dim maxTotalVolume As Double

  '*******************************************************************************
  'Logistics Variables
  '*******************************************************************************
  Dim curOutputRow As Integer
  Dim lastRow As Long
  Dim curRow As Long
  Dim cellRange As String
  Dim vOutputValues As Variant

  '*******************************************************************************
  'define arrays to store Ticker Symbol data which will be written to 
  'spreadsheet after it is all calculated. This will make it run much faster
  '*******************************************************************************
  Dim tickerArr(), yearlyArr(), percentArr(), volumeArr() As Variant
  Dim arrIndex As Long

  Dim WS_Count As Integer
  WS_Count = ActiveWorkbook.Worksheets.Count
  
  '*******************************************************************************
  'Loop through all worksheets and capture / write data for each sheet
  '*******************************************************************************
  For I = 1 To WS_Count
    With ActiveWorkbook.Worksheets(I)
      '***************************************************************************
      'reset all variables for this sheet
      '***************************************************************************
      curTicker = .Cells(2, 1).Value
      curTickerTotalVolume = 0
      curTickerOpen = .Cells(2, 3).Value
      curTickerchange = 0
      curTickerPercentChange = 0
      maxPercentIncreaseTicker = ""
      maxPercentIncrease = 0
      maxPercentDecreaseTicker = ""
      maxPercentDecrease = 0
      maxTotalVolumeTicker = ""
      maxTotalVolume = 0
      curOutputRow = 2
      curRow = 2

      '***************************************************************************
      'define and initialize the array element Index variable
      '***************************************************************************
      arrIndex = 1
      '***************************************************************************
      'initialize arrays with header in the first element
      '***************************************************************************
      tickerArr = VBA.Array("Ticker")
      yearlyArr = VBA.Array("Yearly Change")
      percentArr = VBA.Array("Percent Change")
      volumeArr = VBA.Array("Total Stock Volume")

      '***************************************************************************
      'Determne number of rows on active sheet
      '***************************************************************************
      lastRow = .Cells(Rows.Count, 1).End(xlUp).Row

      '***************************************************************************
      'Provide formatting for cells with % data
      '***************************************************************************
      .Range("K:K").NumberFormat = "0.00%"
      .Range("Q2:Q3").NumberFormat = "0.00%"
  
      '***************************************************************************
      'Walk through and process all rows of data on this sheet
      '***************************************************************************
      For thisRow = curRow To lastRow
        curTickerTotalVolume = curTickerTotalVolume + .Cells(thisRow, 7)

        'Check for new ticker symbol in next row
        If .Cells(thisRow + 1, 1).Value <> .Cells(thisRow, 1).Value Then
          'if there is a new Greatest Total Volume, update appropriate variables
          If curTickerTotalVolume >= maxTotalVolume Then
            maxTotalVolume = curTickerTotalVolume
            maxTotalVolumeTicker = curTicker
          End If 'curTickerTotalVolume >= maxTotalVolume

          'Calculate Yearly and Percent Change for this Ticker Symbol
          curTickerchange = (.Cells(thisRow, 6) - curTickerOpen)
          If curTickerOpen = 0 Then
              curTickerPercentChange = 0
          Else
              curTickerPercentChange = (curTickerchange / curTickerOpen)
          End If 'curTickerOpen = 0

          '***********************************************************************
          'Check for new Percent Increase/Decrease, set colors and update 
          'appropriate variables
          '***********************************************************************
          If curTickerPercentChange >= 0 Then
            'Percent Change is positive or zero
            If curTickerPercentChange > maxPercentIncrease Then
              maxPercentIncrease = curTickerPercentChange
              maxPercentIncreaseTicker = curTicker
            End If 'curTickerPercentChange > maxPercentIncrease
            .Cells(curOutputRow, 10).Interior.ColorIndex = 4
          Else 'curTickerPercentChange >= 0
            'Percent Change is negative
            If curTickerPercentChange < maxPercentDecrease Then
              maxPercentDecrease = curTickerPercentChange
              maxPercentDecreaseTicker = curTicker
            End If 'curTickerPercentChange < maxPercentDecrease
            .Cells(curOutputRow, 10).Interior.ColorIndex = 3
          End If 'curTickerPercentChange >= 0
        
          '***********************************************************************
          'Add current values to the arrays
          '***********************************************************************
          ReDim Preserve tickerArr(arrIndex)
          ReDim Preserve yearlyArr(arrIndex)
          ReDim Preserve percentArr(arrIndex)
          ReDim Preserve volumeArr(arrIndex)
          tickerArr(arrIndex) = curTicker
          yearlyArr(arrIndex) = curTickerchange
          percentArr(arrIndex) = curTickerPercentChange
          volumeArr(arrIndex) = curTickerTotalVolume

          'Update Array Index
          arrIndex = arrIndex + 1

          '***********************************************************************
          'Update/Reset Current Variables
          '***********************************************************************
          curTicker = .Cells(thisRow + 1, 1)
          curOutputRow = curOutputRow + 1
          curTickerTotalVolume = 0
          curTickerOpen = .Cells(thisRow + 1, 3).Value
          curTickerchange = 0
          curTickerPercentChange = 0

        End If 'Cells(thisRow + 1, 1).Value <> Cells(thisRow, 1).Value
  
      Next thisRow ' For thisRow = curRow To lastRow
  
      '***************************************************************************
      'Write all Ticker Symbol Summaries to Spreadsheet (stored in arrays)
      '***************************************************************************
      cellRange = "I1:I" & arrIndex
      .Range(cellRange).Value = Application.Transpose(tickerArr)
      cellRange = "J1:J" & arrIndex
      .Range(cellRange).Value = Application.Transpose(yearlyArr)
      cellRange = "K1:K" & arrIndex
      .Range(cellRange).Value = Application.Transpose(percentArr)
      cellRange = "L1:L" & arrIndex
      .Range(cellRange).Value = Application.Transpose(volumeArr)
      
      '***************************************************************************
      'Write Greatest Increase, Decrease and Volume to spreadsheet
      '***************************************************************************

      '***************************************************************************
      'Write Row 1 values (Column Headers)
      '***************************************************************************
      cellRange = "O1:Q1"
      vOutputValues = VBA.Array("", "Ticker", "Value")
      .Range(cellRange).Value = vOutputValues
      
      '***************************************************************************
      'Write Row 2 values (Greatest % Increase)
      '***************************************************************************
      cellRange = "O2:Q2"
      vOutputValues = VBA.Array("Greatest % Increase", maxPercentIncreaseTicker, maxPercentIncrease)
      .Range(cellRange).Value = vOutputValues
      
      '***************************************************************************
      'Write Row 3 values (Greatest % Decrease)
      '***************************************************************************
      cellRange = "O3:Q3"
      vOutputValues = VBA.Array("Greatest % Increase", maxPercentDecreaseTicker, maxPercentDecrease)
      .Range(cellRange).Value = vOutputValues
      
      '***************************************************************************
      'Write Row 4 values (Greatest Total Volume)
      '***************************************************************************
      cellRange = "O4:Q4"
      vOutputValues = VBA.Array("Greatest % Increase", maxTotalVolumeTicker, maxTotalVolume)
      .Range(cellRange).Value = vOutputValues

    End With

  Next I 'For I = 1 To WS_Count

  '*******************************************************************************
  'Let user know the processing is finished 
  '*******************************************************************************
  MsgBox ("Processing complete")

End Sub

