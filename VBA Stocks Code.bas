Attribute VB_Name = "Module1"
Sub StockInfoOutput()

    Dim CurrentWs As Worksheet
    For Each CurrentWs In Worksheets
  
        Dim ticker_name As String
        ticker_name = " "
        
        Dim total_ticker_vol As Double
        total_ticker_vol = 0
        Dim open_price As Double
        open_price = 0
        
        Dim close_price As Double
        close_price = 0
        
        Dim change_in_price As Double
        change_in_price = 0
        
        Dim change_in_percent As Double
        change_in_percent = 0
        
        Dim max_ticker_name As String
        max_ticker_name = " "
        
        Dim min_ticker_name As String
        min_ticker_name = " "
        
        Dim max_percent As Double
        max_percent = 0
        
        Dim min_percent As Double
        min_percent = 0
        
        Dim max_vol_ticker As String
        max_vol_ticker = " "
        
        Dim max_vol As Double
        max_vol = 0
      
        Dim summary_table_row As Long
        summary_table_row = 2
        
        Dim lastrow As Long
        Dim i As Long
    
        lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
        If need_summary_table_header Then
   
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
        Else
            need_summary_table_header = True
        End If
        
 
        open_price = CurrentWs.Cells(2, 3).Value
        
        For i = 2 To lastrow
      
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                ticker_name = CurrentWs.Cells(i, 1).Value
         
                close_price = CurrentWs.Cells(i, 6).Value
                change_in_price = close_price - open_price
             
                If open_price <> 0 Then
                    change_in_percent = (change_in_price / open_price) * 100
                Else
                    MsgBox ("For " & ticker_name & ", Row " & CStr(i) & ": open price =" & open_price & ". Fix <open> field manually and save the spreadsheet.")
                End If
            
                total_ticker_vol = total_ticker_vol + CurrentWs.Cells(i, 7).Value
              
                
                CurrentWs.Range("I" & summary_table_row).Value = ticker_name
           
                CurrentWs.Range("J" & summary_table_row).Value = change_in_price
              
                If (change_in_price > 0) Then
                    
                    CurrentWs.Range("J" & summary_table_row).Interior.ColorIndex = 4
                ElseIf (change_in_price <= 0) Then
                 
                    CurrentWs.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
                CurrentWs.Range("K" & summary_table_row).Value = (CStr(change_in_percent) & "%")
               
                CurrentWs.Range("L" & summary_table_row).Value = total_ticker_volume
                
                summary_table_row = summary_table_row + 1
                              
        change_in_price = 0
                
                close_price = 0
              
                open_price = CurrentWs.Cells(i + 1, 3).Value
              
      
                If (change_in_percent > max_percent) Then
                    max_percent = change_in_percent
                    max_ticker_name = ticker_name
                ElseIf (change_in_percent < min_percent) Then
                    min_percent = change_in_percent
                    min_ticker_name = ticker_name
                End If
                       
                If (total_ticker_vol > max_vol) Then
                    max_vol = total_ticker_vol
                    max_volume_ticker = ticker_name
                End If
                
                change_in_percent = 0
                total_ticker_vol = 0
            
            Else
              total_ticker_vol = total_ticker_vol + CurrentWs.Cells(i, 7).Value
            End If
      
        Next i
        
End Sub

