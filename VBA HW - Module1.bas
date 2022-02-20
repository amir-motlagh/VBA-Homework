Attribute VB_Name = "Module1"
Sub StockMarket()



Dim Ws As Worksheet
Dim Wb As Workbook

Set Wb = ActiveWorkbook


For Each Ws In Wb.Sheets

    Dim Ticker As String
    Ticker = " "
    
    Dim Ticker_Total As Double
    Ticker_Total = 0
    
    Dim Open_Price As Double
    Open_Price = 0
    
    Dim Close_Price As Double
    Close_Price = 0
    
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    Dim Yearly_Change_Percentage As Double
    Yearly_Change_Percentage = 0
    
    Dim Min_Ticker As String
    Min_Ticker = " "
    
    Dim Max_Ticker As String
    Max_Ticker = " "
    
    Dim Min_Percentage As Double
    Min_Percentage = 0
    
    Dim Max_Percentage As Double
    Max_Percentage = 0
    
    Dim Max_Ticker_Volume As String
    Max_Ticker_Volume = " "
    
    Dim Max_Volume As Double
    Max_Volume = 0
    
    
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

Dim LastRow As Long
LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

Open_Price = Ws.Cells(2, 3).Value


For i = 2 To LastRow

    If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
        Ticker = Ws.Cells(i, 1).Value
    
        Close_Price = Ws.Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price
        
        If Open_Price <> 0 Then
        Yearly_Change_Percentage = (Yearly_Change / Open_Price) * 100
    
    End If
    
    
    
    Ticker_Total = Ws.Cells(i, 7).Value + Ticker_Total
    
    Ws.Range("I" & Summary_Table_Row).Value = Ticker
    
    Ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    
    
    
    If Yearly_Change > 0 Then
        Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
    ElseIf Yearly_Change <= 0 Then
        Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    
    End If

    
    Ws.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_Change_Percentage) & "%")
       
    Ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
    
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    
    Open_Price = Ws.Cells(i + 1, 3).Value
    
    If Yearly_Change_Percentage > Max_Percentage Then
        Max_Percentage = Yearly_Change_Percentage
        Max_Ticker = Ticker
        
    ElseIf Yearly_Change_Percentage < Min_Percentage Then
        Min_Percentage = Yearly_Change_Percentage
        Min_Ticker = Ticker
    
    
    End If
    
    
    If Ticker_Total > Max_Volume Then
       Max_Volume = Ticker_Total
       Max_Ticker_Volume = Ticker
    
    End If
    
    
    Yearly_Change_Percentage = 0
    Ticker_Total = 0
    
    Else
    
        Ticker_Total = Ticker_Total + Ws.Cells(i, 7).Value
        
    End If
    
    
Next i


       
        Ws.Range("P2").Value = Max_Ticker
        Ws.Range("P3").Value = Min_Ticker
        Ws.Range("Q4").Value = Max_Volume
        Ws.Range("O2").Value = "Greatest % Increase"
        Ws.Range("O3").Value = "Greatest % Desrease"
        Ws.Range("O4").Value = "Greatest_Total_Volume"
        Ws.Range("Q2").Value = (CStr(Max_Percentage) & "%")
        Ws.Range("Q3").Value = (CStr(Min_Percentage) & "%")
        
    
    
    
Next Ws
         
    
    
End Sub


