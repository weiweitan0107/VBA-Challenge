Attribute VB_Name = "Module1"
Sub Year2014()

Dim Tickername As String
Dim Openprice As Double
Dim Closeprice As Double
Dim VOLtotal As Double
Dim Summary_Table_Row As Integer
Dim MaxPerIncRow As Integer
Dim MaxPerDecRow As Integer
Dim MaxVolRow As Integer

SheetCount = Application.Sheets.Count
For SheetNum = 1 To SheetCount
Worksheets(SheetNum).Activate


Summary_Table_Row = 2  'since Row 1 for labels only'
VOLtotal = 0

For i = 2 To 100000
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then  'find the first row of the new ticker'
          Tickername = Cells(i, 1).Value
          Openprice = Cells(i + 1, 3).Value
          VOLtotal = VOLtotal + CDbl(Cells(i, 7).Value)
          
          Range("K" & Summary_Table_Row).Value = Tickername
          Range("L" & Summary_Table_Row).Value = Openprice

    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then  'find the last row of the new ticker'
          Closeprice = CDbl(Cells(i, 6).Value)
          VOLtotal = VOLtotal + CDbl(Cells(i, 7).Value)
          
          Range("M" & Summary_Table_Row).Value = Closeprice
          Range("N" & Summary_Table_Row).Value = Closeprice - Openprice
          Range("O" & Summary_Table_Row).Value = FormatPercent((Closeprice - Openprice) / Openprice)
          Range("P" & Summary_Table_Row).Value = VOLtotal
        
          'Set the color by value'
          If Range("N" & Summary_Table_Row).Value < 0 Then
            Range("N" & Summary_Table_Row).Interior.ColorIndex = 3
          Else
            Range("N" & Summary_Table_Row).Interior.ColorIndex = 4
          End If
          
          'get ready for the next new ticker'
          Summary_Table_Row = Summary_Table_Row + 1
          VOLtotal = 0
          Openprice = 0
          Closeprice = 0

    Else
          VOLtotal = VOLtotal + CDbl(Cells(i, 7).Value)
    
    End If

Next i


MaxPerIncRow = 2
MaxPerDecRow = 2
MaxVolRow = 2

For i = 2 To 10000
    If Cells(MaxPerIncRow, "O") < Cells(i + 1, "O") Then
        MaxPerIncRow = i + 1
    End If
    
    If Cells(MaxPerDecRow, "O") > Cells(i + 1, "O") Then
        MaxPerDecRow = i + 1
    End If
            
    If Cells(MaxVolRow, "P") < Cells(i + 1, "P") Then
       MaxVolRow = i + 1
    End If
Next i

Cells(2, "T").Value = Cells(MaxPerIncRow, "K").Value
Cells(2, "U").Value = FormatPercent(Cells(MaxPerIncRow, "O").Value)
Cells(3, "T").Value = Cells(MaxPerDecRow, "K").Value
Cells(3, "U").Value = FormatPercent(Cells(MaxPerDecRow, "O").Value)
Cells(4, "T").Value = Cells(MaxVolRow, "K").Value
Cells(4, "U").Value = Cells(MaxVolRow, "P").Value
Cells(2, "S").Value = "Greatest % Increase"
Cells(3, "S").Value = "Greatest % Decrease"
Cells(4, "S").Value = "Greatest Total Volumn"
Cells(1, "T").Value = "Ticker"
Cells(1, "U").Value = "Value"


Next SheetNum


End Sub

