Attribute VB_Name = "Module2"


Sub challeng2()

Dim row As Long
Dim ws As Worksheet
Dim lastRow As Long, firstRow As Long
For Each ws In ActiveWorkbook.Worksheets

row = Cells(Rows.Count, "A").End(xlUp).row
ActiveSheet.Range("A2:A" & row).AdvancedFilter _
Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("J2"), _
Unique:=True


lastRow = Cells(Rows.Count, "A").End(xlUp).row
lastRowTicker = Cells(Rows.Count, "J").End(xlUp).row

For i = 2 To 500

With Range("A1:G" & lastRow)
    .AutoFilter Field:=1, Criteria1:=Range("J" & i)
    firstRow = .Offset(1, 0).SpecialCells(xlCellTypeVisible).Cells(1, 1).row
    lastFilteredRow = .SpecialCells(xlCellTypeVisible).Cells(.SpecialCells(xlCellTypeVisible).Count).row
    lastFilteredRow1 = lastFilteredRow + firstRow - 2
    
    
    Change = Range("F" & lastFilteredRow1) - Range("C" & firstRow)
    percentchange = (Range("F" & lastFilteredRow1) - Range("C" & firstRow)) / Range("C" & firstRow)
    totalvol = Application.WorksheetFunction.Sum(Range("G" & firstRow & ":G" & lastFilteredRow1))
    
    Range("K" & i).Value = Change
    Range("L" & i).Value = percentchange
    Range("M" & i).Value = totalvol
    
    .AutoFilter ' remove the filter
    
End With
Next i


'Max values

   Dim MaxPercent As Double
    Dim MaxCompany As String
    Dim CurrentPercent As Double
    Dim CurrentCompany As String
       Dim MaxPercentD As Double
    Dim MaxCompanyD As String
    Dim CurrentPercentD As Double
    Dim CurrentCompanyD As String
    Dim MaxVol As Double
    Dim MaxCompanyV As String
    Dim CurrentVol As Double
    Dim CurrentCompanyV As String
    
    
    MaxPercent = -1
    MaxCompany = ""
    
' Loop through all cells in column L
    For Each Cell In Range("L1:L" & Cells(Rows.Count, "L").End(xlUp).row)
        ' Check if the value in the cell is a valid number and is between 0 and 100
        If IsNumeric(Cell.Value) And Cell.Value >= 0 And Cell.Value <= 100 Then
            CurrentPercent = Cell.Value
            CurrentCompany = Cell.Offset(0, -2).Value ' Get the company name from column J
            
            ' Check if the current percent is greater than the current maximum percent
            If CurrentPercent > MaxPercent Then
                MaxPercent = CurrentPercent
                MaxCompany = CurrentCompany
            End If
        End If
    Next Cell
    
        For Each Cell In Range("L1:L" & Cells(Rows.Count, "L").End(xlUp).row)
        ' Check if the value in the cell is a valid number and is between 0 and 100
        If IsNumeric(Cell.Value) And Cell.Value >= -100 And Cell.Value < 0 Then
            CurrentPercentD = Abs(Cell.Value)
            CurrentCompanyD = Cell.Offset(0, -2).Value ' Get the company name from column J
            
            ' Check if the current percent is greater than the current maximum percent
            If CurrentPercentD > MaxPercentD Then
                MaxPercentD = CurrentPercentD
                MaxCompanyD = CurrentCompanyD
            End If
        End If
    Next Cell
    
    For Each Cell In Range("LM1:M" & Cells(Rows.Count, "M").End(xlUp).row)
        ' Check if the value in the cell is a valid number and is between 0 and 100
        If IsNumeric(Cell.Value) Then
            CurrentVol = Cell.Value
            CurrentCompanyV = Cell.Offset(0, -3).Value ' Get the company name from column J
            
            ' Check if the current percent is greater than the current maximum percent
            If CurrentVol > MaxVol Then
                MaxVol = CurrentVol
                MaxCompanyV = CurrentCompanyV
            End If
        End If
    Next Cell
    
    
    Range("R2").Value = MaxPercent
    Range("Q2").Value = MaxCompany

    Range("R3").Value = MaxPercentD
    Range("Q3").Value = MaxCompanyD
    

    Range("R4").Value = MaxVol
    Range("Q4").Value = MaxCompanyV

Next ws


'MsgBox "The first row in the filtered data is " & firstRow & vbCrLf & "The last row in the filtered data is " & lastFilteredRow

End Sub







