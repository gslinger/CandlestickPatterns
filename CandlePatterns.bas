Attribute VB_Name = "CandlePatterns"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''   CandleStick Pattern Detection Based on OHLC Data   ''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''     github.com/gslinger '''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' To use you will need to set wb and ws variables to where your data is
' Change constants ColNo and RowNo to fit your worksheet
' Change the ranges for O/H/L/C data in the CandleScan subroutine
' Run CandleScan, for now it will simply print in the assigned column if it fits criteria
' If two candlestick patterns are detected, for now it will just write them both

''''''''''''''''  To-Do  '''''''''''''''''
'    Add more candle stick patterns
'    Make more user friendly
'    Improve algorithms
'    Add some customization/sensitivity options
'    (maybe) plotting
'    (maybe) testing
'    (maybe) add functions to detect candle color, shape, trend etc.
''''''''''''''''''''''''''''''''''''''''''

' This is an educational project, and i welcome any advice, criticism or contributions!
' My main reference on this project has been: https://www.candlescanner.com/patterns-dictionary/

Option Explicit

Public Const DojiPrecision = 0.05               ' Size of Doji Star
Public Const ColNo = 7                          ' Column Number to Put Results
Public Const RowNo = 1                          ' Row Number to Put Results

Public uClose, uOpen, uHigh, uLow As Variant    ' Array variants to hold OHLC Data  (Faster than reading from Cells?)
Public wb As Workbook, ws As Worksheet          ' Workbook and Worksheet Declarations. If Source and Output sheets differ, will need to make another wb/ws



Sub CandleScan()

Set wb = ActiveWorkbook
Set ws = Worksheets("Data")                     ' Change "Data" to source sheet name

uOpen = Range("B2:B2013").Value2
uHigh = Range("C2:C2013").Value2
uLow = Range("D2:D2013").Value2
uClose = Range("E2:E2013").Value2

Dim i As Long

' +3 is because several patterns require several lags of data
For i = LBound(uClose) + 3 To UBound(uClose)
    Call TestCandles(i)
Next


End Sub


Private Sub TestCandles(i As Long)   'need to add screenupdating = false

    'Declare new variables to make formulas more readable
    Dim O, C, H, L, O1, C1, O2, C2, H1, H2, L1, L2 As Double
    O = uOpen(i, 1)
    O1 = uOpen(i - 1, 1)
    O2 = uOpen(i - 2, 1)
    C = uClose(i, 1)
    C1 = uClose(i - 1, 1)
    C2 = uClose(i - 2, 1)
    H = uHigh(i, 1)
    H1 = uHigh(i - 1, 1)
    H2 = uHigh(i - 2, 1)
    L = uLow(i, 1)
    L1 = uLow(i - 1, 1)
    L2 = uLow(i - 2, 1)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''            Doji              ''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Abs(O - C) <= (H - L) * DojiPrecision Then
        Cells(i + RowNo, ColNo).Value = Cells(i + RowNo, ColNo).Value & "Doji"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''      Bearish Engulfing       ''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If C1 > O1 And O > C And O >= C1 And O1 >= C And (O - C) > (C1 - O1) Then
        Cells(i + RowNo, ColNo).Value = Cells(i + RowNo, ColNo).Value & "Bearish Engulfing"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''      Dark Cloud Cover        ''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If C1 > O1 And (C1 + O1) / 2 > C And O > C And O > C1 And C > O1 And (O - C) / (0.001 + (H - L)) > 0.6 Then
        Cells(i + RowNo, ColNo).Value = Cells(i + RowNo, ColNo).Value & "Dark Cloud Cover"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''     Three Outside Down       ''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If C2 > O2 And O1 > C1 And O1 >= C2 And O2 >= C1 And (O1 - C1) > (C2 - O2) And O > C And C < C1 Then
        Cells(i + RowNo, ColNo).Value = Cells(i + RowNo, ColNo).Value & "Three Outside Down"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''      Evening Star Doji       ''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If C2 > O2 And (C2 - O2) / (0.001 + H2 - L2) > 0.6 And (C2 < O1) And (C1 > O1) And (H1 - L1) > (3 * (C1 - O1)) And _
            O > C And O < O1 Then
        Cells(i + RowNo, ColNo).Value = Cells(i + RowNo, ColNo).Value & "Evening Star Doji"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''        Bearish Harami        ''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If C1 > O1 And O > C And O <= C1 And O1 <= C And (O - C) < (C1 - O1) Then
        Cells(i + RowNo, ColNo).Value = Cells(i + RowNo, ColNo).Value & "Bearish Harami"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''        Bearish Harami        ''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub

