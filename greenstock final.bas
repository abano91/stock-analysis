Attribute VB_Name = "Module1"
Sub macrocheck()
Dim testmessage As String
testmessage = "hello world!"
MsgBox (testmessage)

End Sub
Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalvolume = 0

    Dim startingprice As Double
    Dim endingprice As Double

    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = 2 To RowCount

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalvolume = totalvolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingprice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingprice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalvolume
    Cells(4, 3).Value = (endingprice / startingprice) - 1


End Sub

Sub AllStocksAnalysis()

Range("A1").Value = "all stocks 2018"
'create hearder row
 Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

Dim tickers(12) As String

tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

'3a) initialize variables for the starting price and ending price
Dim startingprice As Single
Dim endingprice As Single

'3b)activate the data worksheet
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4)loop through tickers

For i = 0 To 11
ticker = tickers(i)
totalvolume = (0)

'5)loop through rows in the data

Worksheets("2018").Activate
For j = 2 To RowCount

'5a) find the total voulme for the current ticker.
If Cells(j, 1).Value = ticker Then

totalvolume = totalvolume + Cells(j, 8).Value
End If


'5b)find startingprice for the current ticker.

If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
startingprice = Cells(j, 6).Value
End If


'5c)find ending price for the current ticker.

If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
endingprice = Cells(j, 6).Value
End If

Next j

'6)output the data for the current ticker.


Cells(4 + i, 1).Value = ticker
Cells(4 + 1, 2).Value = totalvolume
Cells(4 + 1, 3).Value = endingprice / startingprice - 1

Next i


End Sub

Sub FormatAllStockAnalysisTable()
'formatting

Range("A3:C3").Font.Italic = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C14").NumberFormat = "0.00%"
Columns("B").AutoFit

If Cells(4, 3) > 0 Then
'color the cell green
Cells(4, 3).Interior.Color = vbGreen
 End If
 
 
If Cells(4, 3) > 0 Then
 'color cells red
 Cells(4, 3).Interior.Color = vbRed
 End If
 
 Else
 'clear the cell color
 Cells(4, 3).Interior.Color = xlNone
 End If
 
 datarowstart = 4
 datarowend = 15
 For i = dataroestart To datarowend
 If Cells(i, 3) > 0 Then
 'change the cell color to green
 Cells(i, 3).Interior.Color = vbGreen
 
 ElseIf Cells(i, 3) < 0 Then
 'change the cell colors to red
Cells(i, 3).Interior.Color = vbRed
Else
'clear the cell color
Cells(i, 3).Interior.Color = xlNone

End If
Next i
 
 Worksheet(yearvalue).Activate


Dim startTime As Single
Dim endTime As Single

yearvlaue = InputBox("what year would you like to input analysis on?")
startTime = Timer
endTime = Timer
MsgBox "this code ran in" & (endTime - startTime) & "seconds for the year" & (yearvalue)

Range("A1").Value = "all stocks(" + yearvalue + ")"


Dim startTime As Single
Dim endTime As Single

End Sub



 Sub Clearworksheet()
 Cells.Clear
 

End Sub


