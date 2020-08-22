Sub VBA_Challenge()

Dim ws As Worksheet
Dim total_open As Double
Dim total_close As Double
Dim Vol As Double
Dim Percentage As Double
Dim Increase As String
Dim Increase_value As Double
Dim Decrease As String
Dim Decrease_value As Double
Dim Name_Vol As String
Dim Big_Vol As Double
Dim ticker_name As String
Dim ticker_num As Integer
Dim New_table As Integer


For Each ws In ThisWorkbook.Worksheets

ws.Activate

ticker_num = 0
New_table = 2
total_open = 0
total_colose = 0
Vol = 0


Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"

Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To last_row
    If i = 2 Then
        total_open = Cells(i, 3).Value
        ticker_name = Cells(i, 1).Value
        Vol = Cells(i, 7).Value
        ticker_num = ticker_num + 1
        
    Else
        If (Cells(i, 1).Value <> ticker_name) Then
            
            Cells(ticker_num + 1, 10).Value = ticker_name
            total_close = Cells(i - 1, 6).Value
            Cells(ticker_num + 1, 11).Value = total_close - total_open
            If (total_close - total_open >= 0) Then
                Cells(ticker_num + 1, 11).Interior.ColorIndex = 4
            Else
                Cells(ticker_num + 1, 11).Interior.ColorIndex = 3
            End If
            
            Cells(ticker_num + 1, 13).Value = Vol
            If total_open <> 0 Then
                Cells(ticker_num + 1, 12).Value = ((total_close - total_open)) / total_open
                Cells(ticker_num + 1, 12).NumberFormat = "0.00%"
            Else
                Cells(ticker_num + 1, 12).Value = 0
                Cells(ticker_num + 1, 12).NumberFormat = "0.00%"
            End If
              
            
            ticker_name = Cells(i, 1).Value
            ticker_num = ticker_num + 1
            total_open = Cells(i, 3).Value
            Vol = Cells(i, 7).Value
        Else
            Vol = Vol + Cells(i, 7).Value
        End If
        
    End If
Next i


Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

For i = 2 To ticker_num
    If (i = 2) Then
        Increase = Cells(i, 10).Value
        Increase_value = Cells(i, 12).Value
        Decrease = Cells(i, 10).Value
        Decrease_value = Cells(i, 12).Value
        Name_Vol = Cells(i, 10).Value
        Big_Vol = Cells(i, 13).Value
    Else
        If (Cells(i, 12).Value > Increase_value) Then
            Increase = Cells(i, 10).Value
            Increase_value = Cells(i, 12).Value
        End If
        If (Cells(i, 12).Value < Decrease_value) Then
            Decrease = Cells(i, 10).Value
            Decrease_value = Cells(i, 12).Value
        End If
        If (Cells(i, 13).Value > Big_Vol) Then
            Name_Vol = Cells(i, 10).Value
            Big_Vol = Cells(i, 13).Value
        End If
    End If
Next i
Cells(2, 16).Value = Increase
Cells(2, 17).Value = Increase_value
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 16).Value = Decrease
Cells(3, 17).Value = Decrease_value
Cells(3, 17).NumberFormat = "0.00%"
Cells(4, 16).Value = Name_Vol
Cells(4, 17).Value = Big_Vol
Range("O2:O4").Interior.ColorIndex = 34
Range("P1:Q1").Interior.ColorIndex = 34
Range("J1:M1").Interior.ColorIndex = 34

Next
End Sub