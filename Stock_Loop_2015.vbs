Attribute VB_Name = "Module2"
Sub Stock_Loop_2015()

MsgBox ("Total Volume")

Dim Ticker_Label As String

Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

Dim Yearly_Change As Double

Dim Percent_Change As Double
    
Dim Stock_Open As Double

Dim Stock_Close As Double
    

Dim Summary_Table_Row As Double
Summary_Table_Row = 2

For i = 2 To 760192

    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        Ticker_Label = Cells(i, 1).Value
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

          Range("I" & Summary_Table_Row).Value = Ticker_Label
          Range("L" & Summary_Table_Row).Value = Total_Stock_Volume


    Stock_Close = Cells(i, 6)
       
        If Stock_Open = 0 Then
            Yearly_Change = 0
            Percent_Change = 0
        Else
            Yearly_Change = Stock_Close - Stock_Open
            Percent_Change = Round((Yearly_Change / Stock_Open), 4)
        
        End If

            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

            Summary_Table_Row = Summary_Table_Row + 1

    ElseIf Cells(i - 1, 1).Value <> Cells(i, 1) Then
         Stock_Open = Cells(i, 3)

    Total_Stock_Volume = 0

    Else
    
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
    
    End If


        Next i

MsgBox ("Postive/Negative Percentages")

For j = 2 To 760192

    If Range("J" & j).Value > 0 Then
        Range("J" & j).Interior.ColorIndex = 4

    ElseIf Range("J" & j).Value < 0 Then
            Range("J" & j).Interior.ColorIndex = 3
        
    End If

        Next j
    
 
End Sub
