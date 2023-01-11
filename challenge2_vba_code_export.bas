Attribute VB_Name = "Module1"
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

Sub challenge():

'Declaring excel worksheet workbook variables
Dim Wsheet As Worksheet
Dim Wbook As Workbook

'Declaring calculation variables
Dim TSymbol As String
Dim YChange As Double
Dim BeginOpenValue As Double
Dim EndCloseValue As Double
Dim PercentChange As Double
Dim TSVolume As LongLong

Dim GIncreaseTSymbol As String
Dim GIncreaseValue As Double

Dim GDecreaseTSymbol As String
Dim GDecreaseValue As Double

Dim GTotalVolumeTSymbol As String
Dim GTotalVolumeValue As Double

' Keep track of the location for each ticker symbol in table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim LastRowNumber As Long


Set Wbook = ActiveWorkbook

'Looping through worksheets in the workbook
For Each Wsheet In Wbook.Sheets
    'MsgBox Wsheet.Name
    
    Summary_Table_Row = 2
          
    Wsheet.Cells(1, 9).Value = "Ticker"
    Wsheet.Cells(1, 10).Value = "Yearly Change"
    Wsheet.Cells(1, 11).Value = "Percent Change"
    Wsheet.Cells(1, 12).Value = "Total Stock Volume"
    
    'To identify row number of last row
    LastRowNumber = Wsheet.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox LastRowNumber
    
    
    YChange = 0
    'To make sure the beginning open value is set
    BeginOpenValue = Wsheet.Cells(2, 3).Value
    EndCloseValue = 0
    PercentChange = 0
    TSVolume = 0

    GIncreaseValue = 0
    GDecreaseValue = 0
    GTSVolumeValue = 0
    
    For i = 2 To LastRowNumber
          ' Check if we are still within the same ticker symbol, if it is not...
          If Wsheet.Cells(i + 1, 1).Value <> Wsheet.Cells(i, 1).Value Then
            TSymbol = Wsheet.Cells(i, 1).Value
            
            TSVolume = TSVolume + Wsheet.Cells(i, 7).Value
            EndCloseValue = Wsheet.Cells(i, 6).Value
            
            YChange = EndCloseValue - BeginOpenValue
            PercentChange = YChange / BeginOpenValue
            
            If GIncreaseValue < PercentChange Then
                GIncreaseTSymbol = TSymbol
                GIncreaseValue = PercentChange
            End If
            
            If GDecreaseValue > PercentChange Then
                GDecreaseTSymbol = TSymbol
                GDecreaseValue = PercentChange
            End If
            
            If GTSVolumeValue < TSVolume Then
                GTSVolumeTSymbol = TSymbol
                GTSVolumeValue = TSVolume
            End If
            
            'Displaying calculated values
            Wsheet.Cells(Summary_Table_Row, 9).Value = TSymbol
            
            Wsheet.Cells(Summary_Table_Row, 10).Value = YChange
            Wsheet.Cells(Summary_Table_Row, 11).Value = PercentChange
            
            
            'Conditional Formatting for Yearly Change/Percent Change Column: Negative,0 is red, 1 and above is Green
            If YChange <= 0 Then
                Wsheet.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf YChange > 0 Then
                Wsheet.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
            If PercentChange <= 0 Then
                Wsheet.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf PercentChange > 0 Then
                Wsheet.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
            'Percentage Formatting
            
            Wsheet.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            Wsheet.Cells(Summary_Table_Row, 12).Value = TSVolume
            
            TSVolume = 0
            BeginOpenValue = Wsheet.Cells(i + 1, 3).Value
            EndCloseValue = Wsheet.Cells(i + 1, 6).Value

            'Increasing summary row by 1
            Summary_Table_Row = Summary_Table_Row + 1
                            
          ' If the cell immediately following a row is the same ticker symbol...
          Else
            TSVolume = TSVolume + Wsheet.Cells(i, 7).Value
            EndCloseValue = Wsheet.Cells(i, 6).Value
          End If
          
    Next i
    
    'Displaying calculated greatest values
    Wsheet.Cells(1, 16).Value = "Ticker"
    Wsheet.Cells(1, 17).Value = "Value"
    
    Wsheet.Cells(2, 15).Value = "Greatest % Increase"
    Wsheet.Cells(2, 16).Value = GIncreaseTSymbol
    Wsheet.Cells(2, 17).Value = GIncreaseValue
    Wsheet.Range("Q2").NumberFormat = "0.00%"
    
    Wsheet.Cells(3, 15).Value = "Greatest % Decrease"
    Wsheet.Cells(3, 16).Value = GDecreaseTSymbol
    Wsheet.Cells(3, 17).Value = GDecreaseValue
    Wsheet.Range("Q3").NumberFormat = "0.00%"
    
    Wsheet.Cells(4, 15).Value = "Greatest Total Volume"
    Wsheet.Cells(4, 16).Value = GTSVolumeTSymbol
    Wsheet.Cells(4, 17).Value = GTSVolumeValue
    
    'Autofiting columns widths for calculated value columns
    Wsheet.Columns("I:Q").AutoFit
       
Next Wsheet



End Sub
