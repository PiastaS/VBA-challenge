Sub StockMarket()

Dim i As Long
Dim j As Long
Dim m As Integer
Dim n As Integer
Dim Max As Double
Dim Min As Double
Dim Max_i As Integer
Dim Min_i As Integer
Dim Ticker_Name As String
Dim Volume_Total As Double
Dim New_Index As Double
Dim Rec_Count  As Double
Dim Year_Start_Index As Double
Dim Year_End_Index As Double
Dim Year_Start_Volume As Double
Dim Year_End_Volume As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim New_Value As Double
Dim Old_Value As Double
Dim Min_Percent As Double
Dim Max_Percent As Double
Dim Max_Vol As Double
Dim Percent_Range As Double
Dim Greatest_Volume As Double
Dim Vol_i As Integer
Dim Lastrow As Long

For Each WS In Worksheets

WS.Range("K1").Value = "Ticker"
WS.Range("L1").Value = "Yearly Change"
WS.Range("M1").Value = "Percent Change"
WS.Range("N1").Value = "TotalStockVolume"
WS.Range("Q1").Value = "Ticker"
WS.Range("Q1").Interior.Color = vbYellow
WS.Range("R1").Value = "Value"
WS.Range("R1").Interior.Color = vbYellow
WS.Range("P3").Value = "Greatest % Increase"
WS.Range("P2").Value = "Greatest % Decrease"
WS.Range("P4").Value = "Greatest Total Volume"
WS.Range("K1").Interior.Color = vbYellow
WS.Range("L1").Interior.Color = vbYellow
WS.Range("M1").Interior.Color = vbYellow
WS.Range("N1").Interior.Color = vbYellow
WS.Range("P2").Interior.Color = vbYellow
WS.Range("P3").Interior.Color = vbYellow
WS.Range("P4").Interior.Color = vbYellow

Lastrow = WS.Range("A" & Rows.Count).End(xlUp).Row

Volume_Total = 0
New_Index = 2
Rec_Count = 0
Year_End_Volume = 0

For i = 2 To Lastrow

    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    
        Ticker_Name = WS.Cells(i, 1).Value
       
        Volume_Total = Volume_Total + WS.Cells(i, 7).Value
        
        WS.Range("K" & New_Index).Value = Ticker_Name
        WS.Range("N" & New_Index).Value = Volume_Total
        Volume_Total = 0
        
        Rec_Count = Rec_Count + 1
        Year_End_Index = i
        Year_Start_Index = Year_End_Index - Rec_Count + 1
        
        New_Value = WS.Range("F" & Year_End_Index).Value
        Old_Value = WS.Range("C" & Year_Start_Index).Value
        
        If Old_Value = 0 Then
             WS.Range("M" & New_Index).Value = 0
        ElseIf (New_Value / Old_Value) = 0 Then
            WS.Range("M" & New_Index).Value = 0
        Else
        Yearly_Change = New_Value - Old_Value
        WS.Range("L" & New_Index).Value = Yearly_Change
        
        If Yearly_Change >= 0 Then
            WS.Range("L" & New_Index).Interior.Color = vbGreen
        Else
            WS.Range("L" & New_Index).Interior.Color = vbRed
        End If
        
        Percent_Change = New_Value / Old_Value - 1
        WS.Range("M" & New_Index).Value = Percent_Change
        WS.Range("M" & New_Index).NumberFormat = "0.00%"
        
        End If
        
        Yearly_Change = 0
        Percent_Change = 0
        Rec_Count = 0
        
        New_Index = New_Index + 1
    
    Else
      Volume_Total = Volume_Total + WS.Cells(i, 7).Value
      Rec_Count = Rec_Count + 1
    End If
Next i

 Max = 0
 Min = 0
 Min_i = 0
 Greatest_Volume = 0
 
For m = 2 To New_Index
    If WS.Range("M" & m).Value < Min Then
        Min = WS.Range("M" & m).Value
        Min_i = m
        WS.Range("Q2").Value = WS.Range("K" & Min_i).Value
        WS.Range("R2").Value = Min
        WS.Range("R2").NumberFormat = "0.00%"
    End If
     If WS.Range("M" & m).Value > Max Then
        Max = WS.Range("M" & m).Value
        Max_i = m
        WS.Range("Q3").Value = WS.Range("K" & Max_i).Value
        WS.Range("R3").Value = Max
        WS.Range("R3").NumberFormat = "0.00%"
    End If
    If WS.Range("N" & m).Value > Greatest_Volume Then
        Greatest_Volume = WS.Range("N" & m).Value
        Vol_i = m
        WS.Range("Q4").Value = WS.Range("K" & Vol_i).Value
        WS.Range("R4").Value = Greatest_Volume
    End If

Next m
Next WS

End Sub
