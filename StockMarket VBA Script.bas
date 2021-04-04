Attribute VB_Name = "Module1"
Sub StockData()
For Each ws In Worksheets

Dim LastRow As Long

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Ticker As String
Dim Yearly_Delta As Double
Dim Percent_Delta As Double
Dim Total_Volumne As Double
Dim Open_Price As Double
Dim Close_Price As Double

Yearly_Delta = 0
Percent_Delta = 0
Total_Volume = 0
Open_Price = 0
Close_Price = 0

Dim Summary_Table As Integer
Summary_Table = 2


ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Delta "
ws.Cells(1, 11).Value = "Percent Delta"
ws.Cells(1, 12).Value = "Total Volume"
ws.Cells(2, 15).Value = "Greatest % Inc"
ws.Cells(3, 15).Value = "Greatest % Dec"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Summary_Table = Summary_Table + 1
    
    Ticker_Symbol = ws.Cells(i, 1).Value
    
    Yearly_Delta = ws.Cells(i, 3).Value - ws.Cells(i, 6).Value
    
    Total_Volume = ws.Cells(i, 7).Value + Cells(i, 7).Value
    
    Close_Price = ws.Cells(i, 6).Value
    
    Open_Price = ws.Cells(2, 3).Value
    
    
    Total_Volume = 0
    
    
    
    
    If Open_Price <> 0 Then
    
    Percent_Delta = (Yearly_Delta / Open_Price)
    Else
    Percent_Delta = 0
    End If
    
    ws.Range("K" & Summary_Table).Value = Percent_Delta
    ws.Range("K" & Summary_Table).Style = "Percent"
    
    Open_Price = ws.Cells(i + 1, 3).Value
    Else
    Total_Volume = Total_Volume + Cells(i, 7).Value
    
    ws.Range("L" & Summary_Table).Value = Total_Volume
     
     ws.Range("I" & Summary_Table).Value = Ticker_Symbol
     
     ws.Range("J" & Summary_Table).Value = Yearly_Delta
     
    End If
    
    
Next i

For i = 2 To LastRow

If ws.Cells(i, 10).Value >= 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
Else
ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i

Dim PercentFinalRow As Long
PercentFinalRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
Dim Percent_Max As Double
Dim Percent_Min As Double
Percent_Max = 0
Percent_Min = 0

For i = 2 To PercentFinalRow

If Percent_Max < ws.Cells(i, 11) Then
Percent_Max = ws.Cells(i, 11).Value

ws.Cells(2, 17).Value = Percent_Max
ws.Cells(2, 17).Style = "Percent"
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

ElseIf Percent_Min > ws.Cells(i, 11) Then
Percent_Min = ws.Cells(i, 11).Value

ws.Cells(3, 17).Value = Percent_Min
ws.Cells(3, 17).Style = "Percent"
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

End If
Next i


Next ws


End Sub

