Sub CreateDiabetesChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Glycèmie_De_Richard_Perreault") ' Ensure this matches your sheet name
    
    ' Delete any existing charts on the worksheet
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    
    ' Create a new chart object at cell M5 with a fixed size
    Dim newChartObj As ChartObject
    Set newChartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("M5").Left, _
        Top:=ws.Range("M5").Top, _
        Width:=500, _
        Height:=300)

    With newChartObj.Chart
        ' Set the chart type to a line chart so that connecting lines are drawn
        .ChartType = xlLine

        ' Force blanks to be interpolated so that if there are gaps the available points are joined
        .DisplayBlanksAs = xlInterpolated

        Dim seriesCount As Integer
        seriesCount = 0

        ' SERIES 1: Before Breakfast – dates from A5:A100 and readings from B5:B100
        If Application.WorksheetFunction.count(ws.Range("B5:B100")) > 0 Then
            seriesCount = seriesCount + 1
            With .SeriesCollection.NewSeries
                .Name = "Glycémie à jeun"
                .XValues = ws.Range("A5:A100")
                .Values = ws.Range("B5:B100")
            End With
        End If

        ' SERIES 2: Before Dinner – dates from A5:A100 and readings from D5:D100 (if any data)
        If Application.WorksheetFunction.count(ws.Range("D5:D100")) > 0 Then
            seriesCount = seriesCount + 1
            With .SeriesCollection.NewSeries
                .Name = "Glycémie avant diner"
                .XValues = ws.Range("A5:A100")
                .Values = ws.Range("D5:D100")
            End With
        End If

        ' SERIES 3: Before Supper – dates from A5:A8 and readings from F5:F8
        If Application.WorksheetFunction.count(ws.Range("F5:F8")) > 0 Then
            seriesCount = seriesCount + 1
            With .SeriesCollection.NewSeries
                .Name = "Glycémie avant souper"
                .XValues = ws.Range("A5:A8")
                .Values = ws.Range("F5:F8")
            End With
        End If

        ' SERIES 4: Before Sleeping – dates from A5:A100 and readings from I5:I100 (if any data)
        If Application.WorksheetFunction.count(ws.Range("I5:I100")) > 0 Then
            seriesCount = seriesCount + 1
            With .SeriesCollection.NewSeries
                .Name = "Glycémie avant Dodo"
                .XValues = ws.Range("A5:A100")
                .Values = ws.Range("I5:I100")
            End With
        End If

        ' Format the X-axis (assumed to be dates)
        With .Axes(xlCategory)
            .TickLabels.Orientation = 45
            .HasTitle = True
            .AxisTitle.Text = "Date"
        End With

        ' Format the Y-axis (glucose readings)
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Glucose"
        End With

        ' Show the legend
        .HasLegend = True

        ' Set distinct line colors for the series that were added
        If seriesCount >= 1 Then .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' Red
        If seriesCount >= 2 Then .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(0, 255, 0) ' Green
        If seriesCount >= 3 Then .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(0, 0, 255) ' Blue
        If seriesCount >= 4 Then .SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(255, 165, 0) ' Orange
    End With
End Sub