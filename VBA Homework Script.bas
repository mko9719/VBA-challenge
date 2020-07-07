Attribute VB_Name = "Module3"
Sub Stock_Data()
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Stock Volume"

Dim Ticker As String
Dim Yearly_Change As Double
Dim BOY_Price As Double
Dim EOY_Price As Double
Dim Num_of_Trading_Days As Integer
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Total_Stock_Volume As LongLong

For i = 2 To 705714
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker
        Total_Stock_Volume = Application.WorksheetFunction.SumIf(Range("A2:G705714"), Ticker, Range("G2:G705714"))
        Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        Num_of_Trading_Days = Application.WorksheetFunction.CountIf(Range("A2:A705714"), Ticker) - 1
        BOY_Price = Cells(i - Num_of_Trading_Days, 3).Value
        EOY_Price = Cells(i, 6).Value
        Yearly_Change = EOY_Price - BOY_Price
        Range("J" & Summary_Table_Row).Value = Yearly_Change
        Range("K" & Summary_Table_Row).Value = (EOY_Price - BOY_Price) / BOY_Price
        If Range("K" & Summary_Table_Row).Value >= 0 Then
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf Range("K" & Summary_Table_Row).Value < 0 Then
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        Else
        End If
        Summary_Table_Row = Summary_Table_Row + 1
    End If
Next i
        
End Sub



