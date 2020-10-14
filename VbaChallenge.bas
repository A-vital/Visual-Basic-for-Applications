Attribute VB_Name = "Module1"
Sub stock()

Dim wb As Workbook
Dim ws As Worksheet

Dim Ticker_name As String
Dim RowCount As Long
Dim end_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim stock_volume As Double
Dim summary_table_row As Long
Dim beg_price As Double

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Set wb = ActiveWorkbook

For Each ws In Worksheets

    Ticker_name = ""
    summary_table_row = 2
    end_price = 0
    yearly_change = 0
    percent_change = 0
    stock_volume = 0
    beg_price = 0

 
 RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
 
 beg_price = ws.Cells(2, 3).Value
 
    For i = 2 To RowCount
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker_name = ws.Cells(i, 1).Value
        
            end_price = ws.Cells(2, 6).Value
        
            yearly_change = end_price - beg_price
        
            If beg_price <> 0 Then
                    percent_change = (yearly_change / beg_price) * 100
            
            End If
            
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            Range("I" & summary_table_row).Value = Ticker_name
            Range("J" & summary_table_row).Value = yearly_change
            
            If (yearly_change > 0) Then
                    Range("J" & summary_table_row).Interior.ColorIndex = 4
            
            ElseIf (yearly_change <= 0) Then
                    Range("J" & summary_table_row).Interior.ColorIndex = 3
            End If
                    
            
            Range("K" & summary_table_row).Value = (CStr(percent_change) & "%")
            
            Range("L" & summary_table_row).Value = stock_volume
            
            summary_table_row = summary_table_row + 1
            
            beg_price = Cells(i + 1, 3).Value
            
            percent_change = 0
            stock_volume = 0
        
                
        
        End If
      
    Next i
 
Next ws
End Sub
