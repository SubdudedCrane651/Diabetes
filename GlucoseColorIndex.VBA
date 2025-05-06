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