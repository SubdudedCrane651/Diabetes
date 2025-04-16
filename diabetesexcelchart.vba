Sub CreateDiabetesChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Glycèmie_De_Richard_Perreault")

    ' Delete existing charts
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    ' Create new chart
    Dim newChartObj As ChartObject
    Set newChartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("M5").Left, _
        Top:=ws.Range("M5").Top, _
        Width:=500, _
        Height:=300)

    With newChartObj.Chart
        .ChartType = xlLine
        .DisplayBlanksAs = xlInterpolated

        Dim lastRow As Integer
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

        ' Get unique dates while limiting to 20 days
        Dim uniqueDates As Object
        Set uniqueDates = CreateObject("Scripting.Dictionary")
        
        Dim i As Integer
        For i = 5 To lastRow
            If Not uniqueDates.Exists(ws.Cells(i, 1).Value) Then
                uniqueDates.Add ws.Cells(i, 1).Value, ws.Cells(i, 1).Value
            End If
            If uniqueDates.count >= 20 Then Exit For
        Next i

        ' Define X-axis range (filtered dates)
        Dim xValuesRange As String
        xValuesRange = "A5:A" & lastRow
        
        ' Ensure each series is plotted
        Dim seriesCount As Integer
        seriesCount = 0

        ' SERIES 1: Before Breakfast
        If Application.WorksheetFunction.count(ws.Range("B5:B" & lastRow)) > 0 Then
            seriesCount = seriesCount + 1
            With .SeriesCollection.NewSeries
                .Name = "Glycémie à jeun"
                .XValues = ws.Range(xValuesRange)
                .Values = ws.Range("B5:B" & lastRow)
            End With
        End If

        ' SERIES 2: Before Dinner
        If Application.WorksheetFunction.count(ws.Range("D5:D" & lastRow)) > 0 Then
            seriesCount = seriesCount + 1
            With .SeriesCollection.NewSeries
                .Name = "Glycémie avant diner"
                .XValues = ws.Range(xValuesRange)
                .Values = ws.Range("D5:D" & lastRow)
            End With
        End If

        ' SERIES 3: Before Supper
        If Application.WorksheetFunction.count(ws.Range("F5:F" & lastRow)) > 0 Then
            seriesCount = seriesCount + 1
            With .SeriesCollection.NewSeries
                .Name = "Glycémie avant souper"
                .XValues = ws.Range(xValuesRange)
                .Values = ws.Range("F5:F" & lastRow)
            End With
        End If

        ' SERIES 4: Before Sleeping
        If Application.WorksheetFunction.count(ws.Range("I5:I" & lastRow)) > 0 Then
            seriesCount = seriesCount + 1
            With .SeriesCollection.NewSeries
                .Name = "Glycémie avant Dodo"
                .XValues = ws.Range(xValuesRange)
                .Values = ws.Range("I5:I" & lastRow)
            End With
        End If

        ' Format axes
        With .Axes(xlCategory)
            .TickLabels.Orientation = 45
            .HasTitle = True
            .AxisTitle.Text = "Date"
        End With

        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Glucose"
        End With

        .HasLegend = True

        ' Assign colors to each series
        If seriesCount >= 1 Then .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' Red
        If seriesCount >= 2 Then .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(0, 255, 0) ' Green
        If seriesCount >= 3 Then .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(0, 0, 255) ' Blue
        If seriesCount >= 4 Then .SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(255, 165, 0) ' Orange
    End With
End Sub