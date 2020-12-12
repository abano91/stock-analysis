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

    Dim startingPrice As Double
    Dim endingPrice As Double

    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = 2 To RowCount

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalvolume = totalvolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalvolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1


End Sub

Sub allstockanalysis()

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

For i = 1 To 10
'a line of codehere will run 10 times

For j = 1 To 20
'a line of code here will run 200 times
Next j
Next i

For i = 0 To 11
ticker = tickers(i)
'do stuff with ticker
Next i

'create a nested for loop that outs 1into the cells of all colums A through J

RowCount = Cells(Rows, Count, "A").End(xlUp).Row
For i = 0 To 11
tickers = tickers(i)
totalvolume = 0

Worksheets(yearvalue).Activate
For j = 2 To RowCount
If Cells(j, i).Value = ticker Then
totalvolume = Totalvoume + Cells(j, 8).Value
End If
If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
startingPrice = Cells(j, 6).Vlaue
End If

End Sub

