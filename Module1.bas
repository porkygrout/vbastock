Attribute VB_Name = "Module1"
Sub Stock_Ticker()
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"


Dim Ticker As String

Dim Yearly As Double
Yearly = 0

Dim Volume As String
Volume = 0

Dim Max_Per As Double
Max_Per = 0

Dim Min_Per As Double
Min_Per = 0

Dim Max_Vol As String
Max_Vol = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    Yearly = Yearly + Cells(i, 6).Value - Cells(i, 3).Value
    Volume = Volume + Cells(i, 7).Value
    Range("I" & Summary_Table_Row).Value = Ticker
    Range("J" & Summary_Table_Row).Value = Yearly
    Range("L" & Summary_Table_Row).Value = Volume
    
    If Yearly <> 0 Then
        Range("K" & Summary_Table_Row).Value = Yearly / (Cells(i, 6).Value - Yearly)
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    Else
        Range("K" & Summary_Table_Row).Value = 0
    End If
    
    Summary_Table_Row = Summary_Table_Row + 1
    Yearly = 0
    Volume = 0
Else
    Yearly = Yearly + Cells(i, 6).Value - Cells(i, 3).Value
    Volume = Volume + Cells(i, 7).Value
End If

Next i

lastrow1 = Cells(Rows.Count, 9).End(xlUp).Row

For J = 2 To lastrow1
If Cells(J, 11).Value > Max_Per Then
    Max_Per = Cells(J, 11).Value
    Cells(2, 16).Value = Cells(J, 9).Value
    Cells(2, 17).Value = Max_Per
    Cells(2, 17).NumberFormat = "0.00%"
End If

If Cells(J, 11).Value < Min_Per Then
    Max_Per = Cells(J, 11).Value
    Cells(3, 16).Value = Cells(J, 9).Value
    Cells(3, 17).Value = Max_Per
    
End If

If Cells(J, 12).Value > Max_Vol Then
    Max_Vol = Cells(J, 12).Value
    Cells(4, 16).Value = Cells(J, 9).Value
    Cells(4, 17).Value = Max_Vol
End If

If Cells(J, 10).Value >= 0 Then
    Cells(J, 10).Interior.ColorIndex = 4
Else
    Cells(J, 10).Interior.ColorIndex = 3

End If

Next J

End Sub
