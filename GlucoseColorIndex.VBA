Sub Macro1()
'
' Macro1 Macro
'

'
    Columns("A:A").Select
    Selection.NumberFormat = "[$-fr-CA]d mmmm, yyyy;@"
End Sub

Sub GlucoseColorIndex()

    'High value
    
    For Each cell In Range("B5:C1000")
        If cell.Value > 10 Then
            cell.Font.Color = vbRed
        End If
    Next cell

    For Each cell In Range("D5:D1000")
        If cell.Value > 10 Then
            cell.Font.Color = vbRed
        End If
    Next cell

    For Each cell In Range("F5:F1000")
        If cell.Value > 10 Then
            cell.Font.Color = vbRed
        End If
    Next cell

    For Each cell In Range("I5:I1000")
        If cell.Value > 10 Then
            cell.Font.Color = vbRed
        End If
    Next cell

    For Each cell In Range("K5:K1000")
        If cell.Value > 10 Then
            cell.Font.Color = vbRed
        End If
    Next cell

    For Each cell In Range("a2:k2")
    On Error Resume Next
        If cell.Value > 10 Then
            cell.Font.Color = vbRed
        End If
    Next cell


    'Normal value
    
    For Each cell In Range("B5:C1000")
        If cell.Value <= 10 And cell.Value >= 3 Then
            cell.Font.ColorIndex = 50
        End If
    Next cell

    For Each cell In Range("D5:D1000")
        If cell.Value <= 10 And cell.Value >= 3 Then
            cell.Font.ColorIndex = 50
        End If
    Next cell

    For Each cell In Range("F5:F1000")
        If cell.Value <= 10 And cell.Value >= 3 Then
            cell.Font.ColorIndex = 50
        End If
    Next cell

    For Each cell In Range("I5:I1000")
        If cell.Value <= 10 And cell.Value >= 3 Then
            cell.Font.ColorIndex = 50
        End If
    Next cell

    For Each cell In Range("K5:K1000")
        If cell.Value <= 10 And cell.Value >= 3 Then
            cell.Font.ColorIndex = 50
        End If
    Next cell

    For Each cell In Range("a2:k2")
        If cell.Value <= 10 And cell.Value >= 3 Then
            cell.Font.ColorIndex = 50
        End If
    Next cell

    'Low value
    
    For Each cell In Range("B5:C1000")
        If cell.Value <= 3.9 Then
            cell.Font.Color = vbBlue
        End If
    Next cell

    For Each cell In Range("D5:D1000")
        If cell.Value <= 3.9 Then
            cell.Font.Color = vbBlue
        End If
    Next cell

    For Each cell In Range("F5:F1000")
        If cell.Value <= 3.9 Then
            cell.Font.Color = vbBlue
        End If
    Next cell

    For Each cell In Range("I5:I1000")
        If cell.Value <= 3.9 Then
            cell.Font.Color = vbBlue
        End If
    Next cell

    For Each cell In Range("K5:K1000")
        If cell.Value <= 3.9 Then
            cell.Font.Color = vbBlue
        End If
    Next cell

    For Each cell In Range("a2:k2")
        If cell.Value <= 3.9 Then
            cell.Font.Color = vbBlue
        End If
    Next cell

End Sub

Sub RemoveImageAndRunPython()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim pythonScript As String

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Glycèmie_De_Richard_Perreault") ' Change sheet name if needed

    ' Loop through all shapes to find and delete the image
    For Each shp In ws.Shapes
        If shp.TopLeftCell.Address = "$K$27" Then
            shp.Delete
        End If
    Next shp

    ' Set the Python script path
    pythonScript = """C:\Users\rchrd\Documents\Richard\Diabetes\Glucose_Chart.py""" ' Update with your actual Python script

    ' Run the Python script
    Shell "python " & pythonScript, vbNormalFocus

    MsgBox "Image deleted & Python script executed!", vbInformation, "Success"
End Sub