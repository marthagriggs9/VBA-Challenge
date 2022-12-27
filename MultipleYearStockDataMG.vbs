Sub thisisatest()

'step 1, name all variables
Dim ticker As String
Dim tickertype As Integer
Dim lastrow As Long
Dim openingprice As Double
Dim closingprice As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim totalstockvolume As Double

'I will activate these after I try this intital loop
Dim greatestpercentincrease As Double
Dim greatestpercentdecrease As Double
Dim greatesttotalvolume As Double
Dim greatestpercentincreaseticker As String
Dim greatestpercentdecreaseticker As String
Dim greatesttotalvolumeticker As String


'step 2, find the last row of the sheet, I will also be changing this later to loop through the whole worksheet
'lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
For Each ws In Worksheets
ws.Activate
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'step 3, add headers to the new columns that need to be created
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'step 4, set all variables to 0
ticker = ""
tickertype = 0
openingprice = 0
yearlychange = 0
percentchange = 0
totalstockvolume = 0
'closingprice = 0 may not need to set this variable

'step 5, loop through the list of tickers
For i = 2 To lastrow

'step 6, get the ticker symbol from the row
ticker = Cells(i, 1).Value

'step 7, get opening price for the ticker
If openingprice = 0 Then
    openingprice = Cells(i, 3).Value
End If

'step 8, Find the total stock volume
totalstockvolume = totalstockvolume + Cells(i, 7).Value

'step 9,when the loop gets to a different ticker this will tell it to change the ticker type
If Cells(i + 1, 1).Value <> ticker Then
    tickertype = tickertype + 1
        Cells(tickertype + 1, 9) = ticker
        
'step 10, get the closing price for the given ticker
closingprice = Cells(i, 6)

'step 11, get yearly change value
yearlychange = closingprice - openingprice

'step 12, put the yearly change value in column J of the worksheet
Cells(tickertype + 1, 10).Value = yearlychange

'step 13, color the yearly change column according to value
If yearlychange > 0 Then
    Cells(tickertype + 1, 10).Interior.ColorIndex = 4
ElseIf yearlychange < 0 Then
    Cells(tickertype + 1, 10).Interior.ColorIndex = 3
End If

'step 14, calculate percentage change
If openingprice = 0 Then
percentchange = 0
Else: percentchange = (yearlychange / openingprice)
End If

'step 15, format percent change value to a percent
Cells(tickertype + 1, 11).Value = Format(percentchange, "Percent")

'step 16, set opening price back to 0
openingprice = 0

'step 17, add total stock volume value to column L
Cells(tickertype + 1, 12).Value = totalstockvolume

'step 18, set total stock volume back to 0
totalstockvolume = 0

End If

Next i


'*BONUS*
'step 1, add columns to display the Greatest % increase and decrease, Greatest Total Volume
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'step 2, get the last row
'lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'step 3, intialize variables and set values of variables to the first row in the list to start
greatestpercentincrease = Cells(2, 11).Value
greatestpercentdecrease = Cells(2, 11).Value
greatesttotalvolume = Cells(2, 12).Value

greatestpercentincreaseticker = Cells(2, 9).Value
greatestpercentdecreaseticker = Cells(2, 9).Value
greatesttotalvolumeticker = Cells(2, 9).Value

'step 4, loop through the list of tickers
For i = 2 To lastrow

'step 5, find the ticker with the greatest % increase
If Cells(i, 11).Value > greatestpercentincrease Then
    greatestpercentincrease = Cells(i, 11).Value
    greatestpercentincreaseticker = Cells(i, 9).Value
End If

'step 6, find the ticker with the greates % decrease
If Cells(i, 11).Value < greatestpercentdecrease Then
    greatestpercentdecrease = Cells(i, 11).Value
    greatestpercentdecreaseticker = Cells(i, 9).Value
End If

'step 7, find the ticker with the greatest total volume
If Cells(i, 12).Value > greatesttotalvolume Then
    greatesttotalvolume = Cells(i, 12).Value
    greatesttotalvolumeticker = Cells(i, 9).Value
End If
Next i

'Step 8, add the greatest percent increase,decrease and total volume to each given column
Range("P2").Value = Format(greatestpercentincreaseticker, "Percent")
Range("Q2").Value = Format(greatestpercentincrease, "Percent")
Range("P3").Value = Format(greatestpercentdecreaseticker, "Percent")
Range("Q3").Value = Format(greatestpercentdecrease, "Percent")
Range("P4").Value = greatesttotalvolumeticker
Range("Q4").Value = greatesttotalvolume

Next ws








End Sub
