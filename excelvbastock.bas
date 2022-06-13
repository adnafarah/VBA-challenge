Attribute VB_Name = "Module2"
Sub stock()

Dim ticker As String
Dim yearlychange As Double
Dim percentchange As Double
Dim totalvol As Double
Dim summarytablerow As Integer
Dim openprice As Double
Dim closeprice As Double
Dim max As Double
Dim min As Double
Dim maxvol As Double
Dim ws As Worksheet


'run code for every worksheet in this workbook
For Each ws In Worksheets

totalvol = 0
summarytablerow = 2

'set last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
        'checking which ticker range we are on
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'populate cell with the ticker name + set the column it will populate
        ticker = Cells(i, 1).Value
        Range("M" & summarytablerow).Value = ticker
        
        'set closeprice as the last closeprice at the end of the year & set the column
        closeprice = Cells(i, 6).Value
        Range("V" & summarytablerow).Value = closeprice
           
        'add up the total volumes for the ticket in question
        totalvol = totalvol + Cells(i, 7).Value
        Range("P" & summarytablerow).Value = totalvol
        
     'move to next row in summary table
     summarytablerow = summarytablerow + 1
     
     'set totals as 0, this also resets it for the next ticker in the next loop
   totalvol = 0

    'otherwise
    Else
        'select the open price from the beginning of the year
        If totalvol = 0 Then
        openprice = Cells(i, 3).Value
        Range("U" & summarytablerow).Value = openprice
        End If
        
    'add up the volumes in each row for the ticker in question
    totalvol = totalvol + Cells(i, 7).Value
    
    'calculate yearlychange by subtracting closeprice from openprice & set the columns
    yearlychange = (Cells(summarytablerow, 22).Value) - (Cells(summarytablerow, 21).Value)
    Range("N" & summarytablerow).Value = yearlychange
    
    'calculate the percentage change (not multiplying by 100 as the cells will be formatted to %)
    percentchange = yearlychange / openprice
    Range("O" & summarytablerow).Value = percentchange
    
    End If
    
    
'set conditions so positive changes are in green cells and negative changes are in red cells
'if change is > 0, fill with green
If Cells(summarytablerow, 14).Value >= 0 Then
    Cells(summarytablerow, 14).Interior.ColorIndex = 4
    
    Else
    'if cell is less than 0, fill with red
    Cells(summarytablerow, 14).Interior.ColorIndex = 3
End If

'move to next row
Next i


'bonus

'select column to find the max/min from
Set Rng = Range("O:O")

'use the vba max function on that column
max = Application.WorksheetFunction.max(Rng)
'select which cell to put max value
Cells(5, 20).Value = max

'use the vba min function on that column
min = Application.WorksheetFunction.min(Rng)
'select which cell to put max value
Cells(6, 20).Value = min

'set column to find max volume from
Set vRng = Range("P:P")
'use the vba max function on that column
maxvol = Application.WorksheetFunction.max(vRng)
'select which cell to put max value
Cells(7, 20).Value = maxvol

'Message box to say which worksheet we are on
MsgBox (ws.Name)


'next worksheet
Next

End Sub




