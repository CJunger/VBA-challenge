Attribute VB_Name = "Module1"
Sub Test()

'Loop through all of the worksheets in the active workbook
For Each ws In Worksheets

    'set initial variable for holding stock ticker
    Dim Ticker As String

    'set initial variable for holding total stock volume per stock ticker
    Dim total_stock_volume As Double
    total_stock_volume = 0

    'Keep track of the location for each unique stock ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'name summary table columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Dim open_price As Double
    Dim close_price As Double
    open_price = Cells(2, 3).Value

    'loop through all stocks
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow

    'check if we are within the same stock, if it is not
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'set ticker symbol
    Ticker = Cells(i, 1).Value

    'print ticker in summary table
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    
    'print stock total volume
    ws.Range("L" & Summary_Table_Row).Value = total_stock_volume + Cells(i, 7).Value
    ws.Range("L" & Summary_Table_Row).NumberFormat = "general"
    'find stock closing for year
    close_price = Cells(i, 6).Value
    
    'Print Yearly Change
    ws.Range("J" & Summary_Table_Row).Value = close_price - open_price
    
       'if yearly positive change cell green and if negative color cell red
        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

        ' Otherwise color it red
        Else
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
   

        End If
    
         'Percent change
         ws.Range("K" & Summary_Table_Row).Value = ((close_price - open_price) / open_price)
         ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
         
         'reset stock total volume
         total_stock_volume = 0
         
         'add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
         
        'find stock opening for year
         open_price = Cells(i + 1, 3).Value
    
    Else
        'add to stock total volume
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
    
    End If

Next i

'name greatest % increase, greatest % decrease and greatest total volume, with ticker name and value

ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

Dim PerChange As Range
Dim tvr As Range
Dim decrease As Double
Dim increase As Double
Dim total_volume As Double
Dim LR As Integer
Dim m As Double

LR = Cells(Rows.Count, 11).End(xlUp).Row
Set PerChange = Range("K2:K" & LR)
Set tvr = Range("L2:L" & LR)
m = Application.WorksheetFunction.Min(PerChange)
ws.Range("p3").Value = m
ws.Range("p3").NumberFormat = "0.00%"

n = Application.WorksheetFunction.Max(PerChange)
ws.Range("p2").Value = n
ws.Range("p2").NumberFormat = "0.00%"

o = Application.WorksheetFunction.Max(tvr)
ws.Range("p4").Value = o
ws.Range("p4").NumberFormat = "0"

Next ws

End Sub




