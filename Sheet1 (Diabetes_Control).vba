'For Excel
Sub Macro1()
'
' Macro1 Macro
'
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"
    Columns("B:B").Select
    Selection.NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
    Columns("E:E").Select
    Selection.NumberFormat = "m/d/yyyy"
    Columns("F:F").Select
    Selection.NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
    Columns("I:I").Select
    Selection.NumberFormat = "m/d/yyyy"
    Columns("J:J").Select
    Selection.NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
End Sub

Sub Glucose_Color()


Dim cell As Range

'High value
'Total Average for the mourning
On Error Resume Next
Range("L1").Interior.Color = vbRed

If Range("B2").Value > 10 Then
 Range("B2").Font.Color = vbRed
End If

If Range("F2").Value > 10 Then
 Range("F2").Font.Color = vbRed
End If

If Range("J2").Value > 10 Then
 Range("J2").Font.Color = vbRed
End If

For Each cell In Range("C5:C100")
 If cell.Value > 10 Then
 cell.Font.Color = vbRed
 End If
Next cell

For Each cell In Range("G5:G100")
 If cell.Value > 10 Then
 cell.Font.Color = vbRed
 End If
Next cell

For Each cell In Range("K5:K100")
 If cell.Value > 10 Then
 cell.Font.Color = vbRed
 End If
Next cell

'Normal value
'Total Average for the Afternoon

Range("L2").Interior.ColorIndex = 50

If Range("B2").Value <= 10 And Range("B2").Value >= 3.9 Then
 Range("B2").Font.ColorIndex = 50
End If

If Range("F2").Value <= 10 And Range("F2").Value >= 3.9 Then
 Range("F2").Font.ColorIndex = 50
End If

If Range("J2").Value <= 10 And Range("J2").Value >= 3.9 Then
 Range("J2").Font.ColorIndex = 50
End If

For Each cell In Range("C5:C100")
 If cell.Value <= 10 And cell.Value >= 3.9 Then
 cell.Font.ColorIndex = 50
 End If
Next cell

For Each cell In Range("G5:G100")
  If cell.Value <= 10 And cell.Value >= 3.9 Then
 cell.Font.ColorIndex = 50
 End If
Next cell

For Each cell In Range("K5:K100")
  If cell.Value <= 10 And cell.Value >= 3.9 Then
 cell.Font.ColorIndex = 50
 End If
Next cell

'Low value
'Total average for the Evening

Range("L3").Interior.Color = vbBlue

If Range("B2").Value <= 3.9 Then
 Range("B2").Font.Color = vbBlue
End If

If Range("F2").Value <= 3.9 Then
 Range("F2").Font.Color = vbBlue
End If

If Range("J2").Value <= 3.9 Then
 Range("J2").Font.Color = vbBlue
End If

For Each cell In Range("C5:C100")
 If cell.Value <= 3.9 Then
 cell.Font.Color = vbBlue
 End If
Next cell

For Each cell In Range("G5:G100")
 If cell.Value <= 3.9 Then
 cell.Font.Color = vbBlue
 End If
Next cell

For Each cell In Range("K5:K100")
 If cell.Value <= 3.9 Then
 cell.Font.Color = vbBlue
 End If
Next cell


End Sub

Sub Macro2()
'
' Macro2 Macro
'

'
    Rows("2:2").Select
    Selection.NumberFormat = "0.0"
End Sub