Attribute VB_Name = "Module1"
'----------------------------------------------------------------------------------------------------
'VBA Homework - The VBA of Wall Street'
'----------------------------------------------------------------------------------------------------
 
'-- code to retreive and print all unique ticker symbols in the destination columns (correct code)
Sub VBAHomework()

For Each ws In Worksheets
'Set destination column names in all sheets
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'--Declare Variables
Dim RowCountSource As Long     '--to count total number of rows in the Source column
Dim tvolume As Double               '--variable to store TotalStockVolume
Dim tname As String                    '--to store ticker in the destination column
Dim var1 As Long                       '--for row
Dim var2 As Long                       '--for row
Dim yropenvalue
Dim yrclosevalue
Dim yearlychange
Dim percentchange As Double

'--Initialize the required variables
tvolume = 0
var1 = 2
var2 = 2
yearlychange = 0
percentchange = 0

'determine the number of rows in the source column
RowCountSource = ws.Cells(Rows.Count, 1).End(xlUp).Row

'--PRINTING DISTINCT TICKER AND CALCULATING TOTAL STOCK VOLUME FOR EACH TICKER(GROUP BY TICKER)

'Loop for each entry in column A
For i = 2 To RowCountSource
tvolume = tvolume + ws.Cells(i, 7).Value

'--Check if the same ticker is being referenced or not
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then

'--if the ticker names are not same , then print the ticker name and total stock volume in the destination column
'--else keep adding the total stock volume for the same ticker names
tname = ws.Cells(i, 1).Value
ws.Cells(var1, 9).Value = tname
ws.Cells(var1, 12).Value = tvolume

'Set tvolume to zero to calculate the sum of next ticker name iteration
tvolume = 0

'-- CALCULATE YEARLY CHANGE
               
yropenvalue = ws.Range("C" & var2)
'MsgBox ("Open Value :" + Str(yropenvalue))
yrclosevalue = ws.Range("F" & i)
'MsgBox ("Close Value :" + Str(yrclosevalue))
yearlychange = yrclosevalue - yropenvalue
'MsgBox ("Change Value :" + Str(yearlychange))
ws.Cells(var1, 10).Value = yearlychange

' Determine Percent Change
If yropenvalue = 0 Then
percentchange = 0
Else
yropenvalue = ws.Range("C" & var2)
percentchange = yearlychange / yropenvalue
ws.Cells(var1, 11).Value = percentchange
End If


'conditional formatting that will highlight positive change in green and negative change in red
If ws.Range("J" & var1).Value >= 0 Then
ws.Range("J" & var1).Interior.ColorIndex = 4
Else
ws.Range("J" & var1).Interior.ColorIndex = 3
End If

'Increment to add a new row for the new ticker name and ticker volume in destination
var1 = var1 + 1
var2 = i + 1

End If
Next i
Next ws

'--SOLUTION FOR CHALLENGES
For Each w2 In Worksheets
'--Set the headers
w2.Cells(2, 14).Value = "Greatest % Increase"
w2.Cells(3, 14).Value = "Greatest % Decrease"
w2.Cells(4, 14).Value = "Greatest Total Volume"
w2.Cells(1, 15).Value = "Ticker"
w2.Cells(1, 16).Value = "Value"
Next w2

'--CALCULATING MAXIMUM STOCK VOLUME
Dim rowcount1 As Long
'--For loop to calculate 'Maximum Stock volume'
For Each ws In Worksheets
'-- Find the number of rows in the destination columns
rowcount1 = ws.Cells(Rows.Count, 9).End(xlUp).Row
'MsgBox ("RowCount for total volume: " + Str(rowcount1))
'--Declare a temporary variable to store max value between iterations
Dim temp
'--Declare a temporary variable to store the row number of the maximum total stock volume to retrieve ticker
Dim var3 As Long
'--Intialize variable
temp = 0
'--Loop to find the maximum total stock volume
For i = 2 To rowcount1
If ws.Cells(i, 12).Value > temp Then
temp = ws.Cells(i, 12).Value
var3 = i
End If
Next i
'--Print the maximum total stock volume and it's corresponding ticker
ws.Cells(4, 16).Value = temp
ws.Cells(4, 15).Value = ws.Cells(var3, 9).Value
'MsgBox (temp)
Next ws

'--CALCULATING GREATEST % INCREASE
Dim rowcount2 As Long
For Each ws In Worksheets
'-- Find the number of rows in the destination columns
rowcount2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
'MsgBox ("RowCount for Greatest Increase: " + Str(rowcount2))
'--Declare a temporary variable to store max value between iterations
Dim temp1
'--Declare a temporary variable to store the row number of the maximum total stock volume to retrieve ticker
Dim var4 As Long
'--Intialize variable
temp1 = 0
var4 = 0
'--For loop to calculate 'Greatest Increase'
For i = 2 To rowcount2
If ws.Cells(i, 11).Value > temp1 Then
temp1 = ws.Cells(i, 11).Value
var4 = i
End If
Next i
'--Print the %greatest increase value and it's corresponding ticker
ws.Cells(2, 16).Value = temp1
ws.Cells(2, 15).Value = ws.Cells(var4, 9).Value
'MsgBox (temp1)
Next ws

'--CALCULATING GREATEST % DECREASE
Dim rowcount3 As Long
For Each ws In Worksheets
'-- Find the number of rows in the destination columns
rowcount3 = ws.Cells(Rows.Count, 9).End(xlUp).Row
'MsgBox ("RowCount for Greatest Decrease: " + Str(rowcount3))
'--Declare a temporary variable to store max value between iterations
Dim temp2
'--Declare a temporary variable to store the row number of the maximum total stock volume to retrieve ticker
Dim var5 As Long
'--Intialize variable
temp2 = 0
'--For loop to calculate 'Greatest Decrease'
For i = 2 To rowcount3
If ws.Cells(i, 11).Value < temp2 Then
temp2 = ws.Cells(i, 11).Value
var5 = i
End If
Next i
'--Print the greatest decrease percent value and it's corresponding ticker
ws.Cells(3, 16).Value = temp2
ws.Cells(3, 15).Value = ws.Cells(var5, 9).Value
'MsgBox (temp2)
Next ws

Dim RowCountDest As Long        '--to store total number of rows in the destination column
For Each ws In Worksheets
'Adding percentage symbol
ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"

'Adding Cell Colors to the Headers
ws.Range("O1").Font.Bold = True
ws.Range("P1").Font.Bold = True
ws.Range("N2").Font.Bold = True
ws.Range("N3").Font.Bold = True
ws.Range("N4").Font.Bold = True

'--Borders for the summary
With ws.Range("N1:P4").Borders
.LineStyle = xlContinuous
.Weight = xlThin
.ColorIndex = 1
End With

'--Borders for the destination column
RowCountDest = ws.Cells(Rows.Count, 9).End(xlUp).Row
'MsgBox (Str(RowCountDest))
With ws.Range(("I1:L") & RowCountDest).Borders
.LineStyle = xlContinuous
.Weight = xlThin
.ColorIndex = 1
End With

'Adding Cell Colors to the Headers
ws.Cells(1, 9).Interior.ColorIndex = 6
ws.Cells(1, 10).Interior.ColorIndex = 6
ws.Cells(1, 11).Interior.ColorIndex = 6
ws.Cells(1, 12).Interior.ColorIndex = 6
'Formatting the text to be bold in destination header cells
ws.Cells(1, 9).Font.Bold = True
ws.Cells(1, 10).Font.Bold = True
ws.Cells(1, 11).Font.Bold = True
ws.Cells(1, 12).Font.Bold = True

Next ws
End Sub



