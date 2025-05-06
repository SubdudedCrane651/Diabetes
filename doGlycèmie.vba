Sub doGlycemie()
    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim countRow As Integer
    Dim currDate As Date
    Dim glucoseSumMorning As Double, glucoseCountMorning As Integer
    Dim glucoseSumLunch As Double, glucoseCountLunch As Integer
    Dim glucoseSumDinner As Double, glucoseCountDinner As Integer
    Dim glucoseSumEvening As Double, glucoseCountEvening As Integer
    Dim TimeVar As Date
    Dim RowOutput As Integer

    ' Set references to sheets
    Set wsInput = ActiveWorkbook.Sheets("Diabetes_Control")
    Set wsOutput = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault")

    ' Clear previous data
    Call GlucoseDelete
    
    'Mourning Reading

    countRow = 5
    RowOutput = 5 ' Start inserting at row 5
    currDate = wsInput.Cells(countRow, 1).Value

    ' Initialize counters
    glucoseSumMorning = 0 : glucoseCountMorning = 0
    glucoseSumLunch = 0 : glucoseCountLunch = 0
    glucoseSumDinner = 0 : glucoseCountDinner = 0
    glucoseSumEvening = 0 : glucoseCountEvening = 0

    While wsInput.Cells(countRow, 1).Value <> ""
        ' Get Date and Time values
        If wsInput.Cells(countRow, 1).Value <> currDate Then
            ' Store averaged results in output sheet
            'wsOutput.Cells(RowOutput, 1).Value = Format(currDate, "MM/DD/YYYY")

            If glucoseCountMorning > 0 Then
                wsOutput.Cells(RowOutput, 2).Value = Round(glucoseSumMorning / glucoseCountMorning, 1) ' 2 AM
                'wsOutput.Cells(RowOutput, 3).Value = "02:00 AM"
            End If
            
            If glucoseCountLunch > 0 Then
                wsOutput.Cells(RowOutput, 4).Value = Round(glucoseSumLunch / glucoseCountLunch, 1) ' Before lunch (9 AM - 1 PM)
                'wsOutput.Cells(RowOutput, 5).Value = "12:00 PM"
            End If
            
            If glucoseCountDinner > 0 Then
                wsOutput.Cells(RowOutput, 6).Value = Round(glucoseSumDinner / glucoseCountDinner, 1) ' Before dinner (1 PM - 7 PM)
                'wsOutput.Cells(RowOutput, 7).Value = "06:00 PM"
            End If
            
            If glucoseCountEvening > 0 Then
                wsOutput.Cells(RowOutput, 9).Value = Round(glucoseSumEvening / glucoseCountEvening, 1) ' Evening (9 PM - 11:59 PM)
                'wsOutput.Cells(RowOutput, 9).Value = "10:00 PM"
            End If
            
            ' Move to next row for output
            RowOutput = RowOutput + 1
            currDate = wsInput.Cells(countRow, 1).Value

            ' Reset counters for next date
            glucoseSumMorning = 0 : glucoseCountMorning = 0
            glucoseSumLunch = 0 : glucoseCountLunch = 0
            glucoseSumDinner = 0 : glucoseCountDinner = 0
            glucoseSumEvening = 0 : glucoseCountEvening = 0
        End If

        TimeVar = wsInput.Cells(countRow, 2).Value

        ' Categorize glucose readings by time range
        If TimeVar >= TimeValue("12:00 AM") And TimeVar < TimeValue("9:00 AM") Then
            glucoseSumMorning = glucoseSumMorning + wsInput.Cells(countRow, 3).Value
            glucoseCountMorning = glucoseCountMorning + 1
        ElseIf TimeVar >= TimeValue("9:00 AM") And TimeVar < TimeValue("1:00 PM") Then
            glucoseSumLunch = glucoseSumLunch + wsInput.Cells(countRow, 7).Value
            glucoseCountLunch = glucoseCountLunch + 1
        ElseIf TimeVar >= TimeValue("1:00 PM") And TimeVar < TimeValue("7:00 PM") Then
            glucoseSumDinner = glucoseSumDinner + wsInput.Cells(countRow, 7).Value
            glucoseCountDinner = glucoseCountDinner + 1
        ElseIf TimeVar >= TimeValue("9:00 PM") And TimeVar <= TimeValue("11:59 PM") Then
            glucoseSumEvening = glucoseSumEvening + wsInput.Cells(countRow, 11).Value
            glucoseCountEvening = glucoseCountEvening + 1
        End If

        countRow = countRow + 1
    Wend
    
    ' Store final day's readings
    'wsOutput.Cells(RowOutput, 1).Value = Format(currDate, "MM/DD/YYYY")
    If glucoseCountMorning > 0 Then wsOutput.Cells(RowOutput, 2).Value = Round(glucoseSumMorning / glucoseCountMorning, 1)
    If glucoseCountLunch > 0 Then wsOutput.Cells(RowOutput, 4).Value = Round(glucoseSumLunch / glucoseCountLunch, 1)
    If glucoseCountDinner > 0 Then wsOutput.Cells(RowOutput, 6).Value = Round(glucoseSumDinner / glucoseCountDinner, 1)
    If glucoseCountEvening > 0 Then wsOutput.Cells(RowOutput, 9).Value = Round(glucoseSumEvening / glucoseCountEvening, 1)
    
    'Dinner and Lunch Reading
    DoIt = True
    countRow = 5
    RowOutput = 5 ' Start inserting at row 5
    currDate = wsInput.Cells(countRow, 5).Value

    ' Initialize counters
    glucoseSumLunch = 0 : glucoseCountLunch = 0
    glucoseSumAfternoon = 0 : glucoseCountAfternoon = 0

    While wsInput.Cells(countRow, 5).Value <> ""
        TimeVar = wsInput.Cells(countRow, 6).Value

        ' If moving to a new date, insert results and reset counters
        If wsInput.Cells(countRow, 1).Value <> currDate And currDate <> 0 Then
            ' Store averaged results in **one row per date**
            'wsOutput.Cells(RowOutput, 1).Value = Format(currDate, "MM/DD/YYYY") ' Date

            If glucoseCountLunch > 0 Then
                'wsOutput.Cells(RowOutput, 2).Value = "12:00 PM" ' Fixed time
                wsOutput.Cells(RowOutput, 4).Value = Round(glucoseSumLunch / glucoseCountLunch, 1) ' Avg lunch reading
            End If

            If glucoseCountAfternoon > 0 Then
                'wsOutput.Cells(RowOutput, 4).Value = "2:00 PM" ' Fixed time
                wsOutput.Cells(RowOutput, 6).Value = Round(glucoseSumAfternoon / glucoseCountAfternoon, 1) ' Avg afternoon reading
                RowOutput = RowOutput - 1
            End If

            RowOutput = RowOutput + 1 ' Move to next row

            ' Reset values for the next date
            glucoseSumLunch = 0 : glucoseCountLunch = 0
            glucoseSumAfternoon = 0 : glucoseCountAfternoon = 0
            currDate = wsInput.Cells(countRow, 5).Value
        End If

        ' Categorize glucose readings into lunch (12 PM) and afternoon (2 PM)
        If TimeVar >= TimeValue("9:00 AM") And TimeVar < TimeValue("1:00 PM") Then
            glucoseSumLunch = glucoseSumLunch + wsInput.Cells(countRow, 7).Value
            glucoseCountLunch = glucoseCountLunch + 1
        ElseIf TimeVar >= TimeValue("1:00 PM") And TimeVar < TimeValue("5:00 PM") Then
            glucoseSumAfternoon = glucoseSumAfternoon + wsInput.Cells(countRow, 7).Value
            glucoseCountAfternoon = glucoseCountAfternoon + 1
        End If
        
        countRow = countRow + 1
        
    Wend

    ' Store final day's readings (to ensure last date is captured)
    If currDate <> 0 Then
        wsOutput.Cells(RowOutput, 1).Value = Format(currDate, "MM/DD/YYYY") ' Date

        If glucoseCountLunch > 0 Then
            'wsOutput.Cells(RowOutput, 2).Value = "12:00 PM" ' Fixed time
            wsOutput.Cells(RowOutput, 4).Value = Round(glucoseSumLunch / glucoseCountLunch, 1) ' Avg lunch reading
        End If

        If glucoseCountAfternoon > 0 Then
            'wsOutput.Cells(RowOutput, 4).Value = "2:00 PM" ' Fixed time
            wsOutput.Cells(RowOutput, 6).Value = Round(glucoseSumAfternoon / glucoseCountAfternoon, 1) ' Avg afternoon reading
        End If
    End If

    
    'Evening Reading
    
    countRow = 5
    RowOutput = 5 ' Start inserting at row 5
    currDate = wsInput.Cells(countRow, 9).Value
    
    ' Reset counters for next date
    glucoseSumMorning = 0 : glucoseCountMorning = 0
    glucoseSumLunch = 0 : glucoseCountLunch = 0
    glucoseSumDinner = 0 : glucoseCountDinner = 0
    glucoseSumEvening = 0 : glucoseCountEvening = 0
    
    While wsInput.Cells(countRow, 9).Value <> ""
        ' Get Date and Time values
        If wsInput.Cells(countRow, 1).Value <> currDate Then
            ' Store averaged results in output sheet
            'wsOutput.Cells(RowOutput, 1).Value = Format(currDate, "MM/DD/YYYY")

            If glucoseCountMorning > 0 Then
                wsOutput.Cells(RowOutput, 2).Value = Round(glucoseSumMorning / glucoseCountMorning, 1) ' 2 AM
                'wsOutput.Cells(RowOutput, 3).Value = "02:00 AM"
            End If
            
            If glucoseCountLunch > 0 Then
                wsOutput.Cells(RowOutput, 4).Value = Round(glucoseSumLunch / glucoseCountLunch, 1) ' Before lunch (9 AM - 1 PM)
                'wsOutput.Cells(RowOutput, 5).Value = "12:00 PM"
            End If
            
            If glucoseCountDinner > 0 Then
                wsOutput.Cells(RowOutput, 6).Value = Round(glucoseSumDinner / glucoseCountDinner, 1) ' Before dinner (1 PM - 7 PM)
                'wsOutput.Cells(RowOutput, 7).Value = "06:00 PM"
            End If
            
            If glucoseCountEvening > 0 Then
                wsOutput.Cells(RowOutput, 9).Value = Round(glucoseSumEvening / glucoseCountEvening, 1) ' Evening (9 PM - 11:59 PM)
                'wsOutput.Cells(RowOutput, 9).Value = "10:00 PM"
            End If
            
            ' Move to next row for output
            RowOutput = RowOutput + 1
            currDate = wsInput.Cells(countRow, 8).Value

            ' Reset counters for next date
            glucoseSumMorning = 0 : glucoseCountMorning = 0
            glucoseSumLunch = 0 : glucoseCountLunch = 0
            glucoseSumDinner = 0 : glucoseCountDinner = 0
            glucoseSumEvening = 0 : glucoseCountEvening = 0
        End If

        TimeVar = wsInput.Cells(countRow, 10).Value

        ' Categorize glucose readings by time range
        If TimeVar >= TimeValue("12:00 AM") And TimeVar < TimeValue("9:00 AM") Then
            glucoseSumMorning = glucoseSumMorning + wsInput.Cells(countRow, 3).Value
            glucoseCountMorning = glucoseCountMorning + 1
        ElseIf TimeVar >= TimeValue("9:00 AM") And TimeVar < TimeValue("1:00 PM") Then
            glucoseSumLunch = glucoseSumLunch + wsInput.Cells(countRow, 7).Value
            glucoseCountLunch = glucoseCountLunch + 1
        ElseIf TimeVar >= TimeValue("1:00 PM") And TimeVar < TimeValue("7:00 PM") Then
            glucoseSumDinner = glucoseSumDinner + wsInput.Cells(countRow, 7).Value
            glucoseCountDinner = glucoseCountDinner + 1
        ElseIf TimeVar >= TimeValue("9:00 PM") And TimeVar <= TimeValue("11:59 PM") Then
            glucoseSumEvening = glucoseSumEvening + wsInput.Cells(countRow, 11).Value
            glucoseCountEvening = glucoseCountEvening + 1
        End If
        
        countRow = countRow + 1
    Wend

    ' Store final day's readings
    'wsOutput.Cells(RowOutput, 1).Value = Format(currDate, "MM/DD/YYYY")
    If glucoseCountMorning > 0 Then wsOutput.Cells(RowOutput, 2).Value = Round(glucoseSumMorning / glucoseCountMorning, 1)
    If glucoseCountLunch > 0 Then wsOutput.Cells(RowOutput, 4).Value = Round(glucoseSumLunch / glucoseCountLunch, 1)
    If glucoseCountDinner > 0 Then wsOutput.Cells(RowOutput, 6).Value = Round(glucoseSumDinner / glucoseCountDinner, 1)
    If glucoseCountEvening > 0 Then wsOutput.Cells(RowOutput, 9).Value = Round(glucoseSumEvening / glucoseCountEvening, 1)
    
    Call GlucoseSort

    Call CalculateRoundedAverageWithCellsFixed

    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 2) = "=ROUND(AVERAGE($B$5:$B$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 4) = "=ROUND(AVERAGE($D$5:$D$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 6) = "=ROUND(AVERAGE($F$5:$F$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 9) = "=ROUND(AVERAGE($I$5:$I$1000),1)"

    Call DeleteRowsWithZeroAverage

    Call GlucoseColorIndex

    Call sheet1.Glucose_Color

    MsgBox "Glucose readings successfully categorized and exported!", vbInformation, "Success"
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