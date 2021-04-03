Attribute VB_Name = "Module1"
Sub Stock_Market_Data()
    
    'Loop through each worksheet in the workbook
    Dim ws As Worksheet
    For Each ws In Worksheets

        'Declare All Variables
        Dim vol_total As Double
        Dim summary_table_row As Integer
        Dim date_open As Double
        Dim date_close As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim lastrow As Long
        
        'Set the starting values
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        vol_total = 0
        summary_table_row = 2
        date_open = 0
        date_close = 0
        yearly_change = 0
        percent_change = 0
        
        
        'Create labels for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"


        'Loop to search through each row, but not for the headers
        For i = 2 To lastrow
            
            'Conditional logic to determine opening value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                date_open = ws.Cells(i, 3).Value
            End If

            'Add to the "Total Stock Value" amount
            vol_total = vol_total + ws.Cells(i, 7)

            'Conditional to determine when the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'Display each ticker value in column I and total volume in column L
                ws.Cells(summary_table_row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summary_table_row, 12).Value = vol_total

                'Calculate and format the yearly change
                date_close = ws.Cells(i, 6).Value
                yearly_change = date_close - date_open
                ws.Cells(summary_table_row, 10).Value = yearly_change
                ws.Cells(summary_table_row, 10).NumberFormat = "$0.00"

                'Conditional formatting to show positive change as green and negative change as red
                If yearly_change >= 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                End If
                
        'Find the percent change
        If date_open = 0 And date_close = 0 Then
            percent_change = 0
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        ElseIf date_open = 0 Then
            'Cannot divide by 0, but needs to reflect the percent change from 0
            Dim new_stock As String
            new_stock = "New Stock"
            ws.Cells(summary_table_row, 11).Value = new_stock
        Else
            percent_change = yearly_change / date_open
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
            End If
            

               'Increment summary_table_row and reset other values
                summary_table_row = summary_table_row + 1
                vol_total = 0
                date_open = 0
                date_close = 0
                yearly_change = 0
                percent_change = 0
                
             
            End If
            
        
        Next i
        
        'Attempt to Calculate Greatest Percentage Increase, Decrease, and greatest total volume.
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    

    Next ws

End Sub

