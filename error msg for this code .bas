Attribute VB_Name = "Module1"
Sub macrocheck()
Dim testmessage As String
testmessage = "hello world!"
MsgBox (testmessage)

End Sub

Sub DQanalysis()

Worksheets("DQAnalysis").Activate

Range("A1").Value = "DAQO (Ticker:DQ)"

'create a header row'

Cells(3, 1).Value = "year"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = " Return"
rowstart = 2
Rowend = 3013
Totalvolume = 0
For i = rowstart To Rowend
'increasetotalvolume"
Totalvolume = Totalvolme + Cells(i, 8).Value
Next i


End Sub
