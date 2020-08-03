Attribute VB_Name = "Module2"
Sub Summary()

'Set a variable for holding the ticker
Dim ticker As String

'Set a variable for holding yearly change
Dim yearly_change As Double
yearly_change = 0

'Set variables for opening values and closing values
Dim open_value As Double
Dim close_value As Double

'Set a variable for holding percent change
Dim percent_change As Double
percent_change = 0

'Set a variable for holding total stock volume
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Set the location of the summary table
Dim Summary_table_row As Integer

'Set the variable number of rows'
Dim numberRows As Long

'Set a variable for the active worksheet
Dim Sh As Worksheet

'Loop the sheets
For Each Sh In ActiveWorkbook.Worksheets

'Reset the rows of the summary table
Summary_table_row = 2

'Print the headers in row 1
Sh.Range("I1") = "Ticker"
Sh.Range("J1") = "Yearly_Change"
Sh.Range("K1") = "Percent_Change"
Sh.Range("L1") = "Total_Stock_Volume"

'Set the variable number of rows
numberRows = activeSheet.UsedRange.Rows.Count

'Set the value of the open_value variable
open_value = Sh.Range("C2")

'Loop the rows

    For i = 2 To numberRows

'Assign a value to the variable 'ticker'
ticker = Sh.Cells(i, 1).Value

'Set a conditional rule for reading the file
If Sh.Cells(i + 1, 1).Value <> ticker Then

'Print the ticker in the summary table
Sh.Range("I" & Summary_table_row).Value = Sh.Cells(i, 1).Value

'Add the Total Stock Volume
Total_Stock_Volume = Total_Stock_Volume + Sh.Cells(i, 7).Value

'Print the total stock volume in the summary table
Sh.Range("L" & Summary_table_row).Value = Total_Stock_Volume

'Calculate the yearly change if the ticker changes
close_value = Sh.Cells(i, 6).Value
yearly_change = close_value - open_value

'Print yearly change in summary table
Sh.Range("J" & Summary_table_row).Value = yearly_change

'Format yearly change output in green and red
If Sh.Range("J" & Summary_table_row).Value > 0 Then
Sh.Range("J" & Summary_table_row).Interior.ColorIndex = 4
Else
Sh.Range("J" & Summary_table_row).Interior.ColorIndex = 3
End If

'Calculate the percent change from the open price to the closing price
If open_value = 0 Then
open_value = open_value + 0.0001
Else
percent_change = yearly_change / open_value
End If

'Print the percent_change in the summary_table
Sh.Range("K" & Summary_table_row).Value = percent_change

'Format the percent_change in percentage
Sh.Range("K" & Summary_table_row).NumberFormat = "0.00%"

'Reset the open value for the next line
open_value = Sh.Cells(i + 1, 3)

'Increment the summary table row to the next line
Summary_table_row = Summary_table_row + 1

'Reset the total stock volume for the next ticker
Total_Stock_Volume = 0

'If the ticker is the same
Else
'Add the Total Stock Volume
Total_Stock_Volume = Total_Stock_Volume + Sh.Cells(i, 7).Value

End If

Next i

'Bonus'
'Calculate the greatest increase, decrease and total volume

'Set the summary table
Sh.Range("P1").Value = "Ticker"
Sh.Range("Q1").Value = "Value"
Sh.Range("O2").Value = "Greateast % Increase"
Sh.Range("O3").Value = "Greatest % Decrease"
Sh.Range("O4").Value = "Greatest Total Volume"

'Set and print the variable Greatest Increase

Dim GreatestIncrease As Double

GreatestIncrease = WorksheetFunction.max(Sh.Range("K:K"))
Sh.Range("Q2").Value = GreatestIncrease
Sh.Range("Q2").NumberFormat = "0.00%"

'Set and print the variable Greatest Decrease

Dim GreatestDecrease As Double

GreatestDecrease = WorksheetFunction.Min(Sh.Range("K:K"))
Sh.Range("Q3").Value = GreatestDecrease
Sh.Range("Q3").NumberFormat = "0.00%"

'Set and print the variable Greatest Stock Volume

Dim GreatestTotalVolume As String

GreatestTotalVolume = WorksheetFunction.max(Sh.Range("L:L"))
Sh.Range("Q4").Value = GreatestTotalVolume

'Set variables to retrieve the Ticker matching each greatest number

Dim MatchingTicker As String
Dim SearchingGreatestpercentage As Double
Dim SearchingGreatestTotalVolume As String

For j = 2 To numberRows

MatchingTicker = Sh.Cells(j, 9).Value
SearchingGreatestpercentage = Sh.Cells(j, 11).Value
SearchingGreatestTotalVolume = Sh.Cells(j, 12).Value

'Retrieve and Print MatchingTicker for the Greatest increase

If SearchingGreatestpercentage = GreatestIncrease Then
    Sh.Range("P2").Value = MatchingTicker
End If

'Retrieve and Print MatchingTicker for the Greatest decrease
If SearchingGreatestpercentage = GreatestDecrease Then
    Sh.Range("P3").Value = MatchingTicker
End If

'Retrieve the matching ticker for the Greatest Total Stock volume

If SearchingGreatestTotalVolume = GreatestTotalVolume Then
    Sh.Range("P4").Value = MatchingTicker
End If

Next j

Next Sh

End Sub

