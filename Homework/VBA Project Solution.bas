Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data():



'' What are my variables/dims?
'Dim ticker As String
'Dim number_tickers As Integer
'Dim yearly_change As Double
'Dim year_open As Double
'Dim year_close As Double
'Dim percent_change As Double
'Dim Total As Double
'Dim total_stock_volume As Double
'Dim greatest_percent_increase As Double
'Dim greatest_percent_decrese As Double
'Dim greatest_total_volume As Double
'Dim Lastrow As Long


'Iterate through All worksheets
For Each ws In ActiveWorkbook.Worksheets

'Row titles/headers for the variables

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"


'RANGE FOR THE "GREATEST % INCREASE"
    Range("P2").Value = "Greatest % Increase"

'RANGE FOR THE "GREATEST % DECREASE"
    Range("P3").Value = "Greatest % Decrease"

'RANGE FOR THE "GREATEST TOTAL VOLUME"
   Range("P4").Value = "Greatest Total Volume"

 ' RANGE FOR THE TICKER VALUE
    Range("Q1").Value = "Ticker"

' RANGE FOR THE VALUE OF STOCK
    Range("R1").Value = "Value"


' ///////////////////////////////////////////////////////////////////


'           //// PART 1 EASY ////////
                    '/// What are my variables/dims?
                        Dim Total As Double
                        Dim ticker As String

                    '//what is the Row number of the last row
                              RowCount = Cells(Rows.Count, 1).End(xlUp).Row

                    '//Row titles/headers for the variables
                        Range("I1").Value = "Ticker"
                        Range("J1").Value = "Total Stock Volume"

            For i = 2 To RowCount

             '//If the ticker is different then print this
                            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                             ' Print Total Ticker results in variable
                            Total = Total + Cells(i, 7).Value

                            Else

                            'Print the ticker symbol as following
                            Cells(i, 9).Value = Total
                            Cells(i, 10).Value = Total

                            ' Total reset amount is
                            Total = 0

            End If
                            Next i

End Sub
