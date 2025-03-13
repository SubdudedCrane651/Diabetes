Sub doGlyc èmie()
    Dim b As Integer
    Dim count As Integer
    Dim count2 As Integer
    Dim count3 As Integer
    Dim Time As Date
    Dim Beforelunch As Boolean
    Dim samedate As Boolean
    Dim dobeforelunch As Boolean

    Sheets("Glycèmie_De_Richard_Perreault").Select
    Call GlucoseDelete
    Beforelunch = False
    samedate = False
    'check fasting And before breakfast
    count3 = 0
    count4 = 0
    breakfast = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 1)
    While breakfast <> Empty
        breakfast = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 1)
        breakfast2 = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1)
        If breakfast = breakfast2 Then
            ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = breakfast2
            
            breakfastglycèmie = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 3)
            Time = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 2)
            If Time > "1:00:00" And Time <= "9:00:00" Then
                ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 2) = breakfastglycèmie
            Else
                ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 4) = breakfastglycèmie
            End If
            count4 = count4 + 1
        Else
            If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = Empty Then
                ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = breakfast
            End If
            count3 = count3 + 1
        End If
    Wend

    'Check before dinner
    count3 = 0
    count4 = 0
    lunch = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 5)
    While lunch <> Empty
        lunch = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 5)
        lunch2 = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1)
        If lunch2 = lunch Then
            ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = lunch2
            lunchglycèmie = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 7)
            Time = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 2)
            ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 6) = lunchglycèmie
            count4 = count4 + 1
        Else
            count3 = count3 + 1
            If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = Empty Then
                ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = lunch
            End If
        End If
    Wend

    'Check before dodo
    count3 = 0
    count4 = 0
    dodo = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 9)
    While dodo <> Empty
        dodo = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 9)
        dodo2 = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1)
        If dodo = dodo2 Then
            ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = dodo2
            dinerglycèmie = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 11)
            Time = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5 + count4, 10)
            If Time > "21:00:00" Then
                ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 9) = dinerglycèmie
            Else
                ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 6) = dinerglycèmie
            End If
            count4 = count4 + 1
        Else
            count3 = count3 + 1
            If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = Empty Then
                ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + count3, 1) = dodo
            End If
        End If
    Wend

    Call GlucoseSort

    Call CalculateRoundedAverageWithCellsFixed
           
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 2) = "=ROUND(AVERAGE($B$5:$B$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 4) = "=ROUND(AVERAGE($DB$5:$D$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 6) = "=ROUND(AVERAGE($F$5:$F$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 9) = "=ROUND(AVERAGE($I$5:$I$1000),1)"
    
    Call DeleteRowsWithZeroAverage
    
    Call GlucoseColorIndex

End Sub

Sub CalculateRoundedAverageWithCellsFixed()
    Dim cellRange As Range
    Dim averageValue As Double
    Dim roundedAverageValue As Double
    Dim daysavg As Integer
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet

    Set sheet1 = ActiveWorkbook.Sheets("Diabetes_Control")
    Set sheet2 = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault")

    For daysavg = 0 To sheet1.Cells(2, 14).Value - 1

        ' Define the range of cells you want to average using Cells
        Set cellRange = sheet2.Range(sheet2.Cells(5 + daysavg, 2), sheet2.Cells(5 + daysavg, 9))
        
        ' Calculate the average of the cell range
        averageValue = 0
        On Error Resume Next
        averageValue = WorksheetFunction.Average(cellRange)
        If IsError(averageValue) Or IsEmpty(averageValue) Then averageValue = 0
        On Error GoTo 0

        ' Round the average value to one decimal place
        roundedAverageValue = WorksheetFunction.Round(averageValue, 1)
        
        ' Assign the rounded average value to the target cell
        sheet2.Cells(5 + daysavg, 11).Value = roundedAverageValue
    Next daysavg
End Sub

Sub DeleteRowsWithZeroAverage()
    Dim sheet2 As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set sheet2 = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault")
    lastRow = sheet2.Cells(sheet2.Rows.count, 11).End(xlUp).Row

    ' Loop through each row from bottom to top
    For i = lastRow To 5 Step - 1
        If sheet2.Cells(i, 11).Value = 0 Then
            sheet2.Rows(i).Delete
        End If
    Next i
End Sub

Sub GlucoseDelete()
    Rows("5:1000").Select
    Selection.ClearContents
    Start = ActiveWorkbook.Sheets("Diabetes_Control").Cells(5, 1)
    For Days = 0 To ActiveWorkbook.Sheets("Diabetes_Control").Cells(2, 14) - 1
        ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(5 + Days, 1) = Start - Days
    Next
End Sub

Sub GlucoseSort()
    Range("A5:J100").Select
    ActiveWorkbook.Worksheets("Glycèmie_De_Richard_Perreault").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Glycèmie_De_Richard_Perreault").Sort.SortFields. _
        Add2 Key : = Range("A5:A100"), SortOn : = xlSortOnValues, Order : = xlDescending, _
        DataOption : = xlSortNormal
    With ActiveWorkbook.Worksheets("Glycèmie_De_Richard_Perreault").Sort
        .SetRange Range("A5:L1000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:A").Select
    Selection.NumberFormat = "[$-fr-CA]d mmmm, yyyy;@"
    Rows("1:1").Select
End Sub