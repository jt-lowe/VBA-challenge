
Sub JL_VBA_Challenge()


'BONUS - Loop through all worksheets

For Each ws in Worksheets
    ws.Select

    'Define boundaries of the data
        'finding last column to determine where we will print the output data
        Dim last_col As Long

        last_col = Range("A1").End(xlToRight).Column

        'finding last row to determine how many times the code will loop
        Dim last_row As Long
        
        last_row = Range("A1").End(xlDown).Row
        
    'Set a variables to hold output values for ticker, yearly change, percent change and total stock volume
        dim output_ticker as Integer
        dim output_yc as integer
        dim output_pc as integer
        dim output_tsv as integer

        output_ticker = last_col + 2
        output_yc =  output_ticker + 1
        output_pc = output_ticker + 2
        output_tsv = output_ticker + 3

    'Print output table headers
        Cells(1, output_ticker).Value = "Ticker"
        Cells(1, output_yc).Value = "Yearly Change"
        Cells(1, output_pc).Value = "Percent Change"
        Cells(1, output_tsv).Value = "Total Stock Volume"

    'Set a variable to hold the ticker symbol, <open> value, <close> value, the yearly change, the % change, the total <vol> for the year
        Dim ticker_symbol As String
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim vol_total As Double

    'Initialize all variables
        open_price = Cells(2, 3).Value 'We know the first open_price value as it is the first in the table
        close_price = 0
        yearly_change = close_price - open_price
        percent_change = 0
        vol_total = 0

    'Set a variable to track which row in output table to work on, starting at row 2 to account for headers
        Dim output_row As Integer
        output_row = 2

    'Begin looping through dataset
        For i = 2 To last_row

        'Check if ticker_symbol is the same, if not then..
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

                'Set variable to read the current ticker_symbol, then print
                    ticker_symbol = Cells(i, 1)
                    Cells(output_row, output_ticker).Value = ticker_symbol

                'Set the new value for "close_price"
                    close_price = Cells(i, 6).Value

                'Calculate "year change", then print into cell
                    yearly_change = close_price - open_price
                    Cells(output_row, output_yc).Value = yearly_change

                'Apply conditional formatting, color value reference http://dmcritchie.mvps.org/excel/colors.htm
                    If yearly_change > 0 Then
                        Cells(output_row, output_yc).Interior.ColorIndex = 4 'Green for positive
                    Else
                        Cells(output_row, output_yc).Interior.ColorIndex = 3 'Red for negative
                    End If

                'As % change is mathematically reliant on open_price not being 0. Check if open_price = 0, if it does not...
                    If open_price <> 0 Then

                    'Calculate % change, then print
                        percent_change = (yearly_change / open_price)
                        Cells(output_row, output_pc).Value = percent_change

                    'Format cell to read as a percentage
                        Cells(output_row, output_pc).NumberFormat = "0.00%"

                    'If open_price = 0, our result is undefined
                    Else
                        Cells(output_row, output_pc).Value = "Undefined"
                    End If

                'Add to vol_total, then print
                    vol_total = vol_total + Cells(i, 7).Value
                    Cells(output_row, output_tsv).Value = vol_total

                'Reset vol_total to 0
                    vol_total = 0

                'Increment "output_row"
                    output_row = output_row + 1

                'Set the new value for "open_price"
                    open_price = Cells(i + 1, 3).Value
            Else
            
            'If the ticker symbols are the same in both rows, we just need to add to vol_total
                vol_total = vol_total + Cells(i, 7).Value
            End If
            
        Next i



'BONUS --------------------------------------------------------------------------------------

    'Set up the table for bonus output, beginning 2 columns after end of last table
        dim bonus_output as Integer

        bonus_output = output_tsv + 2

    'Print out bonus output table
        Cells(1,bonus_output+1).Value = "Ticker"
        Cells(1,bonus_output+2).Value = "Value"
        Cells(2,bonus_output).Value = "Greatest % Increase"
        Cells(3,bonus_output).Value = "Greatest % Decrease"
        Cells(4,bonus_output).Value =  "Greatest Total Volume"

    'Fit column of bonus output contents
        Columns(bonus_output).Autofit

    'Create variable to hold Greatest Increase, Greatest Decrease and Greatest Volume Total, and corresponding ticker
        dim greatest_inc as Double
        dim greatest_dec as Double
        dim greatest_vol as Double
        dim greatest_inc_tic as String
        dim greatest_dec_tic as String
        dim greatest_vol_tic as String

    'Initialise above variables as 0
        greatest_inc = 0
        greatest_dec = 0
        greatest_vol = 0

    'Open loop to check all output values, we can use the previous "output_row" variable to determine how many times the loop will need to run
    For i = 2 to output_row - 1

            'Check if value stored in greatest_inc is less than what is in the output table, change to the greater value, store corresponding ticker symbol
            If Cells(i,output_pc) <> "Undefined" Then
                If greatest_inc < Cells(i,output_pc).Value Then
                    greatest_inc = Cells(i,output_pc).Value
                    greatest_inc_tic = Cells(i,output_ticker).Value
                End If
            End If
           'Check if value stored in greatest_dec is greater than what is in the output table, change to the lower value, store corresponding ticker symbol
            If Cells(i,output_pc) <> "Undefined" Then
                If greatest_dec > Cells(i, output_pc).Value Then
                    greatest_dec = Cells(i, output_pc).Value
                    greatest_dec_tic = Cells(i,output_ticker).Value
                End If            
            End if
            'Check if value stored in greatest_vol is less than what is in the output table, change to greater value, store corresponding ticker symbol
            If greatest_vol < Cells(i, output_tsv).Value Then
                greatest_vol = Cells(i, output_tsv).Value
                greatest_vol_tic = Cells(i,output_ticker).Value
            End If         
    Next i

    'Print values into table
        Cells(2,bonus_output+1).Value = greatest_inc_tic
        Cells(3,bonus_output+1).Value = greatest_dec_tic
        Cells(4,bonus_output+1).Value = greatest_vol_tic
        Cells(2,bonus_output+2).Value = greatest_inc
            Cells(2,bonus_output+2).NumberFormat = "0.00%" 'Format cell to be a percentage
        Cells(3,bonus_output+2).Value = greatest_dec
            Cells(3,bonus_output+2).NumberFormat = "0.00%" 'Format cell to be a percentage
        Cells(4,bonus_output+2).Value = greatest_vol


'BONUS - Move to next worksheet
Next ws

End Sub