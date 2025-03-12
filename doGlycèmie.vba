Sub doGlycèmie()
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

    count = 0
    doit = True

    For k = 5 To 1000
         On Error GoTo 5
        datevar = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k, 1)
        datevar2 = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - 1, 1)
        If datevar = datevar2 And datevar <> Empty Then
            count = count + 1
            doit = True
            GoTo 5
        Else
            doit = False
        End If
        If count > 0 And Not doit Then
            k = k - 1
            For m = 1 To count
                If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 2) <> Empty Then
                    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k, 2) = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 2)
                    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 1) = "DELETE"
                End If
                If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 4) <> Empty Then
                    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k, 4) = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 4)
                    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 1) = "DELETE"
                End If
                If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 7) <> Empty Then
                    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k, 7) = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 7)
                    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 1) = "DELETE"
                End If
                If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 9) <> Empty Then
                    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k, 9) = ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 9)
                    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(k - m, 1) = "DELETE"
                End If
                count = 0: doit = True
            Next m
        End If
5:          Next k

    For l = 5 To 1000
         On Error GoTo 10
        myrange = Range("B" & l & ":I" & l)
        ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(l, 11) = Application.WorksheetFunction.Average(myrange)
        If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(l, 1) = Empty Then
            ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(l, 11).Delete
10:              Exit For
        End If
    Next

    For h = 5 To 1000
        If ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(h, 1) = "DELETE" Then
            Rows(h).EntireRow.Delete
            h = h - 1
        End If
    Next
    
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 2) = "=ROUND(AVERAGE($B$5:$B$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 4) = "=ROUND(AVERAGE($DB$5:$D$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 6) = "=ROUND(AVERAGE($F$5:$F$1000),1)"
    ActiveWorkbook.Sheets("Glycèmie_De_Richard_Perreault").Cells(2, 9) = "=ROUND(AVERAGE($I$5:$I$1000),1)"
    
    Call GlucoseColorIndex

End Sub
