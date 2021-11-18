# VBA-challenge
Monash Data Bootcamp Week 2 Assignment

'Define range using Range("A1").End(xlDown).Row
Sub Wolf()


'Define boundaries of the data
    Dim lastcol As Long
    Dim lastrow As Long
    
    lastrow = Range("A1").End(xlDown).Row
    lastcol = Range("A1").End(xlToRight).Column
    

'Draw up output table
    Dim tablestart as Long
    
    tablestart = lastcol + 2
    
    Cells(1, tablestart).Value = "Ticker"
    Cells(1, tablestart + 1).Value = "Yearly Change"
    Cells(1, tablestart + 2).Value = "Percent Change"
    Cells(1, tablestart + 3).Value = "Total Stock Volume"
    

'Set a varibable to hold the ticker symbol, set as first ticker in range
    Dim ticker As String
    ticker = Cells(2,1).Value
    
'Set a variable to hold
    'the first <open> value
    Dim first as Double

    'last <close> value
    Dim last as Double

    'the yearly change
    dim year_change as Double

    '% change
    dim percent_change as Double

    'Total <vol> for the year
    dim vol_total as Double

'Initialize all variables
    first = Cells(2,3).Value 'We know the first value as it is the first in the table
    last = 0
    year_change = last-start
    percent_change = 0
    vol_total = 0


'Set a variable to track which row in output table to work on, starting at row 2 to account for headers
    Dim output_row As Integer
    output_row = 2

'Loop through table
    For i = 2 to lastrow

    'Check if ticker symbol is the same, if not then..
        If Cells(i,1).Value <> Cells(i+1,1).Value then

        'Change "ticker" variable to new ticker symbol
            ticker = Cells(i,1)
  
        'Print ticker symbol into the table
            Cells(output_row,tablestart).Value = ticker

        'Set the new value for "last"
            last=Cells(i,6).Value

        'Calculate "year change", then print into cell
            year_change = last-first

        'Print "year_change" value
            Cells(output_row,tablestart+1).Value=year_change

        'Apply conditional formatting, color value from http://dmcritchie.mvps.org/excel/colors.htm
            if year_change > 0 then
                Cells(output_row,tablestart+1).Interior.ColorIndex = 4
            Else
                Cells(output_row,tablestart+1).Interior.ColorIndex = 3
            End If

        'Calculate % change, then print
            percent_change = ((last-first)/first)
            Cells(output_row,tablestart+2).Value = percent_change

        'Add to vol_total, then print
            vol_total = vol_totalt + cells(i,7).Value
            Cells(output_row,tablestart+3).Value = vol_total


        'Increment "output_row"
            output_row = output_row + 1

        Else
           
        'Add to vol_total
            vol_total = vol_totalt +cells(i,7).Value

        End IF
        
        Next i

'Find <open> at beginning of year and <close> at end of year and sum together



'Ensure second column has conditional formatting - Range(range).Interior.ColorIndex = x        x= color value from http://dmcritchie.mvps.org/excel/colors.htm

'Calculate % change

           'Sum all <vol> for ticker
        



End Sub