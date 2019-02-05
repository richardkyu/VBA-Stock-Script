Sub stockCalc()

'Fixed an iteration error at j_a
'Added a way to apply to all sheets.

Dim ws As Worksheet

    For Each ws In Sheets
        ws.Activate
        
'Find the last piece of data in a column
tickerLastRow = Cells(Rows.Count, 1).End(xlUp).Row

'MsgBox (tickerLastRow)

'Find the last cell in row 1

dataLastColumn = Cells(1, Columns.Count).End(xlToLeft).Column

'MsgBox (dataLastColumn)

'Set headers

Cells(1, dataLastColumn + 2) = "Ticker"
Cells(1, dataLastColumn + 3) = "Yearly Change"
Cells(1, dataLastColumn + 4) = "Percent Change"
Cells(1, dataLastColumn + 5) = "Total Stock Volume"
Cells(1, dataLastColumn + 9) = "Ticker"
Cells(1, dataLastColumn + 10) = "Value"


'Set variables for the iteration
Dim i As Long
Dim j As Long
Dim j_a As Long

'Setting up for the hard part
Dim yearBiggest As Double
yearBiggest = 0
Dim yearSmallest As Double
yearSmallest = 0

Dim biggestVolume As Double
biggestVolume = 0


j_a = 2

'Are you kidding me?
Dim totalVol As Double

'You'll need this i later to go down in the sheet when the macro outputs
i = 0

    For j = 2 To tickerLastRow
        
        
        While Cells(j, 1) = Cells(j + 1, 1)
        
        totalVol = totalVol + Cells(j, dataLastColumn)
        
        
        j = j + 1
        
        Wend
        
        If Cells(j, 1) <> Cells(j + 1, 1) Then
        
        'Show ticker difference: MsgBox (Cells((j + 1), 1))

        totalVol = totalVol + (Cells(j, dataLastColumn))
        
        'MsgBox (Cells(j, 1))
        'MsgBox (totalVol)
        'To check sums MsgBox (totalVol)
        
        Cells(2 + i, dataLastColumn + 5) = totalVol
        
        'To label the Stock Ticker
        Cells(2 + i, dataLastColumn + 2) = Cells(j, 1)
        'MsgBox (Cells(j, 1))
        
        'For Yearly Change, msgbox to see what is sub
        'MsgBox (Cells(j_a, 3))
        'MsgBox (Cells(j, 6))
        
        If (Cells(j, 6) < Cells(j_a, 3)) Then
        Cells(2 + i, dataLastColumn + 3) = Abs(Cells(j, 6) - Cells(j_a, 3)) * (-1)
                
        'For Percent Change, if condition to prevent division by zero (stock is worthless)
        If Cells(j_a, 3) = 0 Then
        Cells(2 + i, dataLastColumn + 4) = 0
        
        Else:
        Cells(2 + i, dataLastColumn + 4) = FormatPercent(Cells(2 + i, dataLastColumn + 3) / Cells(j_a, 3), 2, vbTrue, vbTrue)
        End If
            If yearSmallest > Cells(2 + i, dataLastColumn + 3) Then
            
                yearSmallest = Cells(2 + i, dataLastColumn + 3)
                'Print the ticker
                Cells(3, dataLastColumn + 9) = Cells(2 + i, dataLastColumn + 2)
                'Find the largest percentage
                Cells(3, dataLastColumn + 10) = Cells(2 + i, dataLastColumn + 4)
            End If
            
        
        Else: Cells(2 + i, dataLastColumn + 3) = Cells(j, 6) - Cells(j_a, 3)
                
        'For Percent Change, if condition to prevent division by zero (stock is worthless)
       'You need this in both the outer if/else for the calculation to work as the values are generated.
        If Cells(j_a, 3) = 0 Then
        Cells(2 + i, dataLastColumn + 4) = 0
        Else:
        Cells(2 + i, dataLastColumn + 4) = FormatPercent(Cells(2 + i, dataLastColumn + 3) / Cells(j_a, 3), 2, vbTrue, vbTrue)
        End If
        
            If yearBiggest < Cells(2 + i, dataLastColumn + 3) Then
            
                yearBiggest = Cells(2 + i, dataLastColumn + 3)
                'MsgBox (Cells(2 + i, dataLastColumn + 4))
                'Print the ticker
                Cells(2, dataLastColumn + 9) = Cells(2 + i, dataLastColumn + 2)
                'Find the largest percentage
                Cells(2, dataLastColumn + 10) = Cells(2 + i, dataLastColumn + 4)
                
                End If
        
        End If
        
        
        
        
        'To indicate color
        If Cells(2 + i, dataLastColumn + 3) > 0 Then
        Cells(2 + i, dataLastColumn + 3).Interior.ColorIndex = 4
        
        Else: Cells(2 + i, dataLastColumn + 3).Interior.ColorIndex = 3
        End If
        
        
        'To change the j_a with the ticker, you need to increment j_a.
        j_a = j + 1
        
        i = i + 1
        
        'To find the biggest volume
        If biggestVolume < totalVol Then
        Cells(4, dataLastColumn + 9) = Cells(j, 1)
        biggestVolume = totalVol
        Cells(4, dataLastColumn + 10) = biggestVolume
        
        End If
        
        totalVol = 0

        
        End If
        
        
        
        Next j
        



Cells(2, dataLastColumn + 8) = "Greatest % Increase"
Cells(3, dataLastColumn + 8) = "Greatest % Decrease"
Cells(4, dataLastColumn + 8) = "Greatest Total Volume"

Cells(2, dataLastColumn + 10) = Format(Cells(2, dataLastColumn + 10), "Percent")
Cells(3, dataLastColumn + 10) = FormatPercent(Cells(3, dataLastColumn + 10), 2, vbTrue, vbTrue)



'AutoFit All Columns on Worksheet
  ws.Cells.EntireColumn.AutoFit

MsgBox ("Function applied to: " + ws.Name)

        Next ws

      End Sub





Sub clearResults()
 
 For Each ws In Sheets
        ws.Activate
        
        Dim lastRow As Long
        Dim dataLastColumn As Long
        
        
        	dataLastColumn = Cells(1, 1).End(xlToRight).Column
        
       'MsgBox (dataLastColumn) check if output is expected

lastRow = Cells(Rows.Count, dataLastColumn + 2).End(xlUp).Row
            
        'MsgBox (lastRow) check output
        
        Range(Cells(1, dataLastColumn + 2), Cells(lastRow, dataLastColumn + 10)).Clear
        
        
        Next ws

MsgBox ("All Worksheets Reset")

End Sub
















