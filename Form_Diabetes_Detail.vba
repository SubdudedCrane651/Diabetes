'For Access
Option Compare Database

Private Sub Command12_Click()

On Error GoTo Err_Command12_Click

    Call createQry

    Dim stDocName As String

    stDocName = "DiabetesPerMonthChart"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click

End Sub

Private Sub createQry()
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim newSQL As String
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "AveragePerMonth"
    On Error GoTo 0
    
    year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        newSQL = "SELECT Format([datevar],""mm-yyyy"") AS Expr1, Round(Avg(Diabetes.Reading),1) AS AveragePerMonth FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""yyyy"") ORDER BY Format([datevar],""yyyy"");"
    Else
        newSQL = "SELECT Format([datevar],""mm-yyyy"") AS Expr1, Round(Avg(Diabetes.Reading),1) AS AveragePerMonth FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""yyyy"") HAVING (((Format([Datevar], ""yyyy"")) = " & year1 & ")) ORDER BY Format([datevar],""yyyy"");"
    End If
    Set qdf = db.CreateQueryDef("AveragePerMonth", newSQL)
End Sub

Private Sub createQry2()
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim newSQL As String
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "Diabetes_Query"
    On Error GoTo 0
    
    year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        newSQL = "SELECT Diabetes.id, Diabetes.Datevar, Diabetes.Timevar, Diabetes.Food, Diabetes.Reading, Diabetes.Exercise FROM Diabetes ORDER BY Diabetes.Datevar DESC , Diabetes.Timevar;"
    Else
        newSQL = "SELECT Diabetes.id, Diabetes.Datevar, Diabetes.Timevar, Diabetes.Food, Diabetes.Reading, Diabetes.Exercise, Year(Diabetes.Datevar) AS Expr1 FROM Diabetes WHERE (((Year(Diabetes.Datevar)) = " & year1 & ")) ORDER BY Diabetes.Datevar DESC , Diabetes.Timevar;"
    End If
    Set qdf = db.CreateQueryDef("Diabetes_Query", newSQL)
End Sub

Private Sub CreateQry3()

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim newSQL As String

 On Error Resume Next
    DoCmd.DeleteObject acQuery, "AverageControl1"
    On Error GoTo 0
    
    year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        newSQL = "SELECT Format([datevar],""dd-mm-yyyy"") AS Expr1, Round(Avg(diabetes.Reading),1) AS AveragePerMonth FROM diabetes GROUP BY Format([datevar],""dd-mm-yyyy"") ORDER BY Format([datevar],""dd-mm-yyyy"");"
    Else
        newSQL = "SELECT Format([datevar],""dd-mm-yyyy"") AS Expr1, Round(Avg(diabetes.Reading),1) AS AveragePerMonth FROM diabetes WHERE (((diabetes.Datevar) >= #1/1/" & year1 & "# And (diabetes.Datevar) <= #12/31/" & year1 & "#)) GROUP BY Format([datevar],""dd-mm-yyyy"") ORDER BY Format([datevar],""dd-mm-yyyy"");"
    End If
    Set qdf = db.CreateQueryDef("AverageControl1", newSQL)

End Sub

Private Sub Command13_Click()
Dim script
Call CreateQry3
Call Shell("C:/Users/rchrd/AppData/Local/Programs/Python/Python312/python.exe C:\Users\rchrd\Documents\Richard\Diabetes\Richards_Health.py", vbNormalFocus)
End Sub

Private Sub Command26_Click()

On Error GoTo Err_Command26_Click

    Call createQry

    Dim stDocName As String

    stDocName = "Blood_Pressure_Detail"
    DoCmd.OpenForm stDocName, View:=acNormal, DataMode:=acFormPropertySettings, WindowMode:=acWindowNormal

Exit_Command26_Click:
    Exit Sub

Err_Command26_Click:
    MsgBox Err.Description
    Resume Exit_Command26_Click
End Sub

Private Sub Command40_Click()
On Error GoTo Err_Command12_Click
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim SQL As String
    Dim rst As DAO.Recordset
    Dim days As Integer
    Dim daysstr As String
    Dim defaultDays As Integer

    Dim stDocName As String
    
    On Error GoTo Cancel:
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset("Days", dbOpenDynaset)
    
    ' Get the default value from the first record
    If Not rst.EOF Then
        rst.MoveFirst
        defaultDays = rst!days
    Else
        defaultDays = 7 ' Fallback default if the table is empty
    End If
    
    rst.Close
    Set rst = Nothing
    
    days = InputBox("How Many Days back?", "Days", defaultDays)
    
    '' Open table again to update the first record instead of adding a new one
    'Set rst = db.OpenRecordset("Days", dbOpenDynaset)
    
    'If Not rst.EOF Then
        'rst.MoveFirst
        'rst.Edit
        'rst!days = days
        'rst.Update
    'Else
        '' If table is empty, add the record
        'rst.AddNew
        'rst!days = days
        'rst.Update
    'End If
     
    CreateDiabetesQrys (days)
    
    SQL = "SELECT Diabetes.Datevar, Diabetes.Timevar, Diabetes.Food, Diabetes.Reading, Diabetes.Exercise FROM Diabetes WHERE (((Diabetes.Datevar)>=Date()-" + Str(days) + "));"
    
    DoCmd.DeleteObject acQuery, "Readings for Kim"
    On Error GoTo 0
    
    Set qdf = db.CreateQueryDef("Readings for Kim", SQL)

    stDocName = "Diabetes_Detail for Kim"
    DoCmd.OpenReport stDocName, acPreview
    
    'rst.Close
    'Set rst = Nothing
    Set db = Nothing

Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    
Cancel:
    'Cancel

End Sub

Private Sub Command41_Click()
    Dim started As Single: started = Timer
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim days As Integer
    Dim daysstr As String
    Dim defaultDays As Integer
    
    On Error GoTo Cancel
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset("Days", dbOpenDynaset)
    
    ' Get the default value from the first record
    If Not rst.EOF Then
        rst.MoveFirst
        defaultDays = rst!days
    Else
        defaultDays = 7 ' Fallback default if the table is empty
    End If
    
    rst.Close
    Set rst = Nothing
    
    ' Display input box with default value from table
    days = InputBox("How Many Days back?", "Days", defaultDays)
    daysstr = Str(days)
    
    ' Open table again to update the first record instead of adding a new one
    Set rst = db.OpenRecordset("Days", dbOpenDynaset)
    
    If Not rst.EOF Then
        rst.MoveFirst
        rst.Edit
        rst!days = days
        rst.Update
    Else
        ' If table is empty, add the record
        rst.AddNew
        rst!days = days
        rst.Update
    End If
    
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    
 CreateDiabetesQrys (days)

 daysstr = Str(days)
'Call Shell("python3 C:\Users\rchrd\Documents\Richard\Diabetes.py " + Days, vbNormalFocus)
'DoCmd.OpenForm "DiabetesText_form", acNormal, , , acReadOnly, _
 '    , daysstr
'Do: DoEvents: Loop Until Timer - started >= 60
    'Call Open_Excel
    'MsgBox ("Diabetes results ready to be viewed")
    
    'AddDatatToExcel using vba
    AddDataToExcel (days)
    
    'AddDatatToExcel using python
    'Call Shell("C:/Users/rchrd/AppData/Local/Programs/Python/Python312/python.exe C:\Users\rchrd\Documents\Richard\Diabetes\Diabetes_xlsm.py " + Str(days), vbNormalFocus)
    
    
Cancel:
    'Cancel
End Sub




Private Sub Open_Excel()
Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.Workbook
    
    Path = "C:\Users\rchrd\Documents\Richard\Diabetes.xlsm"

    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(Path)
    appExcel.Visible = True
    
    'Do Something or Just Leave Open
    
    Set appExcel = Nothing
    Set myWorkbook = Nothing

End Sub

Private Sub Command46_Click()
Call ImportCSVToAccess
End Sub

Private Sub Datevar_AfterUpdate()
Dim i As Integer
For i = List19.ListCount - 1 To 0 Step -1
    List19.RemoveItem (i)
Next i
Call SortData
Call Populate_Years
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
Dim i As Integer
For i = List19.ListCount - 1 To 0 Step -1
    List19.RemoveItem (i)
Next i
Call SortData
Call Populate_Years
End Sub

Private Sub Form_AfterInsert()
Dim i As Integer
For i = List19.ListCount - 1 To 0 Step -1
    List19.RemoveItem (i)
Next i
Call SortData
Call Populate_Years
End Sub

Private Sub Form_Current()

End Sub

Private Sub Form_Delete(Cancel As Integer)
Dim i As Integer
For i = List19.ListCount - 1 To 0 Step -1
    List19.RemoveItem (i)
Next i
'Call SortData
'Call Populate_Years
End Sub

Private Sub Form_Load()
Me.Form.Caption = First_Name + "'s Diabetes"
Call Sort_Enter
Dim i As Integer
For i = List19.ListCount - 1 To 0 Step -1
    List19.RemoveItem (i)
Next i
Call Populate_Years

Call Calculate_Avg
End Sub
Function First_Name() As String
Dim SQL As String
Dim FirstName As String
Dim count As Integer

'Give First Name

SQL = "SELECT * From Info"

Set rst = CurrentDb.OpenRecordset(SQL)

count = 0
Do Until rst.EOF
FirstName = rst!FirstName
rst.MoveNext
count = count + 1
Loop
rst.Close
First_Name = FirstName

End Function

Private Sub Insert_Click()

Dim currDateTime As Date
Dim LDate As Date

currDateTime = Now()

On Error GoTo Err_Insert_Click

    DoCmd.GoToRecord , , acNewRec
    LDate = DateAdd("yyyy", 3, currDateTime)
    Datevar.Value = Format$(Now(), "Short Date")
    Timevar.Value = Format$(Now(), "HH:MM:SS")
    
    
Exit_Insert_Click:
    Exit Sub

Err_Insert_Click:
    MsgBox Err.Description
    Resume Exit_Insert_Click
    
End Sub
Private Sub Command11_Click()
On Error GoTo Err_Command11_Click

    Call createQry2

    Dim stDocName As String

    stDocName = "Diabetes_Detail"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command11_Click:
    Exit Sub

Err_Command11_Click:
    MsgBox Err.Description
    Resume Exit_Command11_Click
    
End Sub

Private Sub SortData()
Dim TestDB As DAO.Database
Dim rs As DAO.Recordset
Dim SQLString As String
Set rst = CurrentDb.OpenRecordset("SELECT * FROM Diabetes")
rst.Sort = "[Datevar] DESC, [Timevar] ASC"
Set rss = rst.OpenRecordset
Form.Requery
rst.Close
Set rs = Nothing
End Sub

Private Sub SortData2()
Set rst = CurrentDb.OpenRecordset("Select * From Diabetes WHERE YEAR(Datevar) = 2020")
rst.Sort = "[Datevar] DESC, [Timevar] ASC"
Set rss = rst.OpenRecordset
rst.Close
' Recordset rss now contains the sorted table
End Sub

Private Sub List19_Click()

Call Sort_Enter
Call Calculate_Avg

End Sub

Private Sub Sort_Enter()
Dim strTask As String
year1 = List19.Value
If IsNull(year1) Or year1 = "All" Then
strTask = "Select * From Diabetes ORDER BY Datevar DESC,Timevar ASC"
Else
strTask = "Select * From Diabetes WHERE YEAR(Datevar) = " & year1 & " ORDER BY Datevar DESC,Timevar ASC"
End If
Me.RecordSource = strTask
Me.OrderBy = "Datevar DESC,Timevar ASC"
Form.Requery
End Sub

Sub Calculate_Avg()
Dim TestDB As DAO.Database
Dim rs As DAO.Recordset
Dim SQL As String

Call AverageMonths

year1 = List19.Value

If IsNull(year1) Or year1 = "All" Then
    SQL = "SELECT Round(Avg(Diabetes.Reading),1) AS AveragePerYear FROM Diabetes;"
Else
    SQL = "SELECT Format([datevar],""yyyy"") AS Expr1, Round(AVG(Diabetes.Reading),1) AS AveragePerYear FROM Diabetes GROUP BY Format([datevar],""yyyy""), Format([datevar],""yyyy"") HAVING (((Format([Datevar], ""yyyy"")) = " & year1 & ")) ORDER BY Format([datevar],""yyyy"");"
End If

Set rst = CurrentDb.OpenRecordset(SQL)

Dim count As Integer
Do Until rst.EOF
Label22.Caption = rst!AveragePerYear
rst.MoveNext
count = count + 1
Loop

End Sub


Sub Populate_Years()
Dim TestDB As DAO.Database
Dim rs As DAO.Recordset
Dim SQLString As String

Set rst = CurrentDb.OpenRecordset("Select YEAR(Datevar) AS Year From Diabetes ORDER BY Datevar DESC")

Dim count As Integer
Dim year1(1 To 2) As String

year1(1) = ""
year1(2) = ""

List19.AddItem ("All")
Do Until rst.EOF

year1(1) = rst!Year
If year1(1) <> year1(2) Then
'List19.AddItem (year1)
year3 = year1(1)
List19.AddItem (year3)
End If
year1(2) = year1(1)
rst.MoveNext

count = count + 1
Loop

End Sub



Private Sub Command24_Click()
On Error GoTo Err_Command24_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command24_Click:
    Exit Sub

Err_Command24_Click:
    MsgBox Err.Description
    Resume Exit_Command24_Click
    
End Sub
Private Sub Command25_Click()
On Error GoTo Err_Command25_Click


    Screen.PreviousControl.SetFocus
    DoCmd.RunCommand acCmdAutoDial

Exit_Command25_Click:
    Exit Sub

Err_Command25_Click:
    Resume Next
    Resume Exit_Command25_Click
    
End Sub

Private Sub Detail_Paint()
If Me.Reading > 10# Then
    Me.Reading.ForeColor = vbRed
Else
If Me.Reading <= 3.9 Then
    Me.Reading.ForeColor = vbBlue
Else
    Me.Reading.ForeColor = vbBlack
End If
End If
End Sub

Function AverageMonths() As String
Dim SQL As String
Dim doMonth As Boolean

doMonth = True

'January

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgJan FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 01)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgJan FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '01-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgJan <> "" Then
    LblJan.Caption = "Jan " + Str(rst!AvgJan)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblJan.Caption = "Jan 0"
End If

doMonth = True

rst.Close
    
'February

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgFeb FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 02)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgFeb FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '02-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgFeb <> "" Then
    LblFeb.Caption = "Feb " + Str(rst!AvgFeb)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblFeb.Caption = "Feb 0"
End If

doMonth = True

rst.Close

'March

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgMar FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 03)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgMar FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '03-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgMar <> "" Then
    LblMar.Caption = "Mar " + Str(rst!AvgMar)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblMar.Caption = "Mar 0"
End If

doMonth = True

rst.Close

'April

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgApr FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 04)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgApr FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '04-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgApr <> "" Then
    LblApr.Caption = "Apr " + Str(rst!AvgApr)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblApr.Caption = "Apr 0"
End If

doMonth = True

rst.Close

'May

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgMay FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 05)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgMay FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '05-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgMay <> "" Then
    LblMay.Caption = "May " + Str(rst!AvgMay)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblMay.Caption = "May 0"
End If

doMonth = True

rst.Close

'June

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgJun FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 06)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgJun FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '06-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgJun <> "" Then
    LblJun.Caption = "Jun " + Str(rst!AvgJun)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblJun.Caption = "Jun 0"
End If

doMonth = True

rst.Close

'Jully

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgJul FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 07)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgJul FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '07-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgJul <> "" Then
    LblJul.Caption = "Jul " + Str(rst!AvgJul)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblJul.Caption = "Jul 0"
End If

doMonth = True

rst.Close

'August

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgAug FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 08)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgAug FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '08-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgAug <> "" Then
    LblAug.Caption = "Aug " + Str(rst!AvgAug)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblAug.Caption = "Aug 0"
End If

doMonth = True

rst.Close

'September

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgSep FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 09)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgSep FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '09-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgSep <> "" Then
    LblSep.Caption = "Sep " + Str(rst!AvgSep)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblSep.Caption = "Sep 0"
End If

doMonth = True

rst.Close

'October

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgOct FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 10)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgOct FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '10-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgOct <> "" Then
    LblOct.Caption = "Oct " + Str(rst!AvgOct)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblOct.Caption = "Oct 0"
End If

doMonth = True

rst.Close

'November

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgNov FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 11)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgNov FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '11-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgNov <> "" Then
    LblNov.Caption = "Nov " + Str(rst!AvgNov)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblNov.Caption = "Nov 0"
End If

doMonth = True

rst.Close

'December

year1 = List19.Value
    If IsNull(year1) Or year1 = "All" Then
        SQL = "SELECT Format([datevar],""mm"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgDec FROM Diabetes GROUP BY Format([datevar],""mm""), Format([datevar],""mm"") HAVING (((Format([Datevar], ""mm"")) = 12)) ORDER BY Format([datevar],""mm"");"
    Else
        SQL = "SELECT Format([datevar],""mm-yyyy"") AS Monthvar, Round(Avg(Diabetes.Reading),1) AS AvgDec FROM Diabetes GROUP BY Format([datevar],""mm-yyyy""), Format([datevar],""mm-yyyy"") HAVING (((Format([Datevar], ""mm-yyyy"")) = '12-" & year1 & "')) ORDER BY Format([datevar],""mm-yyyy"");"
    End If
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
Do Until rst.EOF
If rst!AvgDec <> "" Then
    LblDec.Caption = "Dec " + Str(rst!AvgDec)
    doMonth = False
Else
    
End If
rst.MoveNext
Loop

If doMonth Then
    LblDec.Caption = "Dec 0"
End If

doMonth = True

rst.Close

End Function
Private Sub Command39_Click()
On Error GoTo Err_Command39_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command39_Click:
    Exit Sub

Err_Command39_Click:
    MsgBox Err.Description
    Resume Exit_Command39_Click
    
End Sub
