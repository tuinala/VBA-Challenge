# VBA-Challenge
Module 2 Assignment VBA Challenge
Sub Stock_Analysis()

For Each ws In Worksheets

    'Set column headers'
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    'Set variables
        Dim Ticker_Name As String
        Dim Stock_Open As Double
            Stock_Open = 0
        Dim Stock_Close As Double
            Stock_Close = 0
        Dim Ticker_TotalVolume As Double
            Ticker_TotalVolume = 0
        Dim Yearly_Change As Double
            Yearly_Change = 0
        Dim Percent_Change As Double
            Percent_Change = 0
        Dim Greatest_Increase As String
        Dim Greatest_Increase_Percent As Double
            Greatest_Increase_Percent = 0
        Dim Greatest_Decrease As String
        Dim Greatest_Decrease_Percent As Double
            Greatest_Decrease_Percent = 0
        Dim Greatest_Volume As String
        Dim Greatest_Volume_Total As Double
            Greatest_Volume_Total = 0
        Dim Lastrow As Long
        Dim Summary_TableRow As Double
            Summary_TableRow = 2
                       
    'Set stock open value'
        Stock_Open = ws.Cells(2, 3).Value
        
    'Set last row of worksheet
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through the worksheets
        For i = 2 To Lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set the starting point for ticker name
        Ticker_Name = ws.Cells(i, 1).Value
    
    'Add ticker total volume
        Ticker_TotalVolume = Ticker_TotalVolume + ws.Cells(i, 7).Value

    'Calculate stock change and percent
        Stock_Close = ws.Cells(i, 6).Value
        Yearly_Change = Stock_Close - Stock_Open
           
    If Stock_Open = 0 Then
        
        Yearly_Change = 0
        
        Percent_Change = 0
    
            Else
        
                Percent_Change = Yearly_Change / Stock_Open
    
End If
    
    'Assign Summary Table headings
        ws.Range("I" & Summary_TableRow).Value = Ticker_Name
        ws.Range("J" & Summary_TableRow).Value = Yearly_Change
        ws.Range("K" & Summary_TableRow).Value = Percent_Change
        ws.Range("K" & Summary_TableRow).NumberFormat = "0.00%"
        ws.Range("L" & Summary_TableRow).Value = Ticker_TotalVolume
    
    'Assign positive and neqative conditional formatting to column
    
    If (Yearly_Change >= 0) Then
    
        ws.Range("J" & Summary_TableRow).Interior.ColorIndex = 4
            
            Else
            
                ws.Range("J" & Summary_TableRow).Interior.ColorIndex = 3

End If
        
    'Set Summary Table for next stock open value
        Summary_TableRow = Summary_TableRow + 1
        Stock_Open = ws.Cells(i + 1, 3).Value
    
    'Set and calculate greatest values
    
    If (Percent_Change > Greatest_Increase_Percent) Then
    
        Greatest_Increase_Percent = Percent_Change
        
        Greatest_Increase = Ticker_Name

ElseIf (Percent_Change < Greatest_Decrease_Percent) Then
    
        Greatest_Decrease_Percent = Percent_Change
        
        Greatest_Decrease = Ticker_Name
    
End If

    If (Ticker_TotalVolume > Greatest_Volume_Total) Then
    
        Greatest_Volume_Total = Ticker_TotalVolume
        
        Greatest_Volume = Ticker_Name
        
End If

    'Trigger Reset values
        Percent_Change = 0
        Ticker_TotalVolume = 0
    
    
    'Get the next ticker total volume
Else: Ticker_TotalVolume = Ticker_TotalVolume + ws.Cells(i, 7).Value
          
End If

Next i
        
    'Assign values to corresponding cells
        ws.Range("P2").Value = Greatest_Increase
        ws.Range("Q2").Value = Greatest_Increase_Percent
        ws.Range("P3").Value = Greatest_Decrease
        ws.Range("Q3").Value = Greatest_Decrease_Percent
        ws.Range("P4").Value = Greatest_Volume
        ws.Range("Q4").Value = Greatest_Volume_Total
    
    'Assign number format
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
    
    'Adjust columns
        ws.Columns("I:Q").AutoFit
        
    'Assign conditional formatting to cells
        
    
    Next ws
    
    End Sub

