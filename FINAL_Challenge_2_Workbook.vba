VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_challenge():

' loop through all worksheets
For Each ws In Worksheets

' add in my side chart headers
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Yearly Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

' define variables
Dim ticker_name As String
Dim i As Long
Dim open_price As Double
Dim ticker_volume As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim ticker_row As Long
Dim last_row As Long

' set the starting values for count
ticker_volume = 0
yearly_change = 0
ticker_row = 2
Start = 2
last_row = ws.Range("A" & Rows.Count).End(xlUp).Row


' count to the last row in column A
For i = 2 To last_row

    ' if a row in the ticker column doesn't match the row below it then...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' defined the ticker name
        ticker_name = ws.Cells(i, 1).Value
        
        ' adding up the final volume for the bottommost row per this ticker name
        ticker_volume = ticker_volume + ws.Cells(i, 7).Value
        
        ' setting the opening price
        open_price = ws.Cells(ticker_row, 3).Value
        
        'getting yearly and % change
        yearly_change = (ws.Cells(i, 6).Value - open_price)
        percent_change = (yearly_change / open_price)

        ' fill out side chart
        ws.Range("J" & ticker_row).Value = ticker_name
        ws.Range("K" & ticker_row).Value = yearly_change
        ws.Range("K" & ticker_row).NumberFormat = "0.00"
        ws.Range("L" & ticker_row).Value = percent_change
        ws.Range("L" & ticker_row).NumberFormat = "0.00%"
        ws.Range("M" & ticker_row).Value = ticker_volume
    
        ' reset
        ticker_volume = 0
        ticker_row = ticker_row + 1
        yearly_change = 0
        open_price = ws.Cells(ticker_row, 3).Value
        
Else
        ticker_volume = ticker_volume + ws.Cells(i, 7).Value
        
    End If
    
    Next i
    
' setting the conditional formatting
 Dim newlastrow As Long
 newlastrow = ws.Range("J" & Rows.Count).End(xlUp).Row
 
 For i = 2 To newlastrow
 
    If ws.Cells(i, 11).Value >= 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
    
    Else
     
     ws.Cells(i, 11).Interior.ColorIndex = 3
    
End If

Next i
    
' getting the % increase and decrease
Dim plus As Double
Dim minus As Double
Dim lr As Long

plus = 0
minus = 0
lr = ws.Range("L" & Rows.Count).End(xlUp).Row

For i = 2 To lr

    If plus < ws.Cells(i, 12).Value Then
    plus = ws.Cells(i, 12).Value
    ws.Range("Q2").Value = plus
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P2").Value = ws.Cells(i, 10).Value
    
    ElseIf minus > ws.Cells(i, 12).Value Then
    minus = ws.Cells(i, 12).Value
    ws.Range("Q3").Value = minus
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P3").Value = ws.Cells(i, 10).Value
    
    End If
    
    Next i
' getting the greatest volume
Dim greatest_vol As Double
Dim volumelastrow As Long

greatest_vol = 0
volumelastrow = ws.Range("M" & Rows.Count).End(xlUp).Row

For i = 2 To volumelastrow

    If greatest_vol < ws.Cells(i, 13).Value Then
    greatest_vol = ws.Cells(i, 13).Value
    ws.Range("Q4") = greatest_vol
    ws.Range("P4").Value = ws.Cells(i, 10).Value
    
    End If
    
    Next i

Next ws

End Sub


