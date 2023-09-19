Attribute VB_Name = "Module1"
Sub credit_card()
Dim Ticker As String
Dim Year_Open As Double
Dim Year_Close As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double

Dim Start_Data As Integer
Dim ws As Worksheet
Dim previous_i As Integer


For Each ws In Worksheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    Start_Data = 2
    previous_i = 1
    Total_Stock_Volume = 0
    
    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 2 To EndRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                previous_i = previous_i + 1
            
                Year_Open = ws.Cells(previous_i, 3).Value
                Year_Close = ws.Cells(i, 6).Value
            
            For j = previous_i To i
            
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
                
            Next j
            
            If Year_Open = 0 Then
                
                Percent_Change = Year_Close
                
            Else
                Yearly_Change = Year_Close - Year_Open
                Percent_Change = Yearly_Change / Year_Open
                
            End If
            
            
            ws.Cells(Start_Data, 9).Value = Ticker
            ws.Cells(Start_Data, 10).Value = Yearly_Change
            ws.Cells(Start_Data, 11).Value = Percent_Change
            
            ws.Cells(Start_Data, 11).NumberFormat = "0.00%"
            ws.Cells(Start_Data, 12).Value = Total_Stock_Volume
            
            Start_Data = Start_Data + 1
            
            
            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0
            
            previous_i = 1
            
        End If
        
    Next i
    
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
    Increase = 0
    Decrease = 0
    Greatest = 0
    
        For k = 3 To kEndRow
        
            last_k = k - 1
            
            current_k = ws.Cells(k, 11).Value
            
            previous_k = ws.Cells(last_k, 11).Value
            
            volume = ws.Cells(k, 12).Value
            
            previous_vol = ws.Cells(last_k, 12).Value
            
            If Increase > current_k And Increase > previous_k Then
                Increase = Increase
                
            ElseIf current_k > Increase And current_k > previous_k Then
                Increase = current_k
                
                increase_name = ws.Cells(k, 9).Value
                
            ElseIf previous_k > Increase And previous_k > current_k Then
            
                Increase = previous_k
                
                increase_name = ws.Cells(last_k, 9).Value
            
            End If
            
            
            
            If Decrease < current_k And Decrease < previous_k Then
                Decrease = current_k
                
                decrease_name = ws.Cells(k, 9).Value
                
            ElseIf previous_k < Increase And previous_k < current_k Then
                
                Decrease = previous_k
                
                decrease_name = ws.Cells(last_k, 9).Value
                
            End If
            
        If Greatest > volume And Greatest > previous_vol Then
            
            Greatest = Greatest
            
            greatest_name = ws.Cells(k, 9).Value
            
        ElseIf volume > Greatest And volume > previous_vol Then
        
            Greatest = previous_vol
            
            greatest_name = ws.Cells(last_k, 9).Value
            
        End If
        
    Next k
    
    
ws.Range("N1").Value = "Column Name"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker Name"
ws.Range("P1").Value = "Value"


ws.Range("O2").Value = increase_name
ws.Range("O3").Value = greatest_name
ws.Range("P2").Value = Increase
ws.Range("P3").Value = Greatest

ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"


jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


    For j = 2 To jEndRow
        
        If ws.Cells(j, 10) > 0 Then
        
            ws.Cells(j, 10).Interior.ColorIndex = 4
            
            Else
            
            ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j
        
Next ws
        
    

End Sub
