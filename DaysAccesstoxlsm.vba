'found in the Forms_Diabetes_Detal form in Access
Private Sub Command41_Click()
    Dim started As Single : started = Timer
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
        defaultDays = rst !days
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
        rst !days = days
        rst.Update
    Else
        ' If table is empty, add the record
        rst.AddNew
        rst !days = days
        rst.Update
    End If
    
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    
    CreateDiabetesQrys(days)

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
    
    
    Cancel :
    'Cancel
End Sub

Public Sub CreateDiabetesQrys(days)
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim newSQL As String
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "Mourning_Gluclose_Reading"
    On Error GoTo 0

    newSQL = "SELECT Format(Diabetes.Datevar,'MM/DD/YYYY') as Datevar, Diabetes.Timevar, Diabetes.Reading FROM Diabetes WHERE (((Diabetes.Datevar) >= Date() - " + Str(days) + ") And ((Diabetes.Timevar) <= #11:00:00 AM#) And ((Year([Diabetes].[Datevar])) = 2025)) ORDER BY Diabetes.Datevar DESC;"
 
    Set qdf = db.CreateQueryDef("Mourning_Gluclose_Reading", newSQL)
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "Afternoon_Glucose_Reading"
    On Error GoTo 0

    newSQL = "SELECT Format(Diabetes.Datevar,'MM/DD/YYYY') as Datevar, Diabetes.Timevar, Diabetes.Reading FROM Diabetes WHERE (((Diabetes.Datevar) >= Date() - " + Str(days) + ") And ((Diabetes.Timevar) >=#11:01:00 AM# And (Diabetes.Timevar)<=#9:00:00 PM#) And ((Year([Diabetes].[Datevar])) = 2025)) ORDER BY Diabetes.Datevar DESC;"
 
    Set qdf = db.CreateQueryDef("Afternoon_Glucose_Reading", newSQL)
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "Evening_Gluclose_Reading"
    On Error GoTo 0

    newSQL = "SELECT Format(Diabetes.Datevar,'MM/DD/YYYY') as Datevar, Diabetes.Timevar, Diabetes.Reading FROM Diabetes WHERE (((Diabetes.Datevar) >= Date() - " + Str(days) + ") And ((Diabetes.Timevar) > #9:01:00 PM#) And ((Year([Diabetes].[Datevar])) = 2025)) ORDER BY Diabetes.Datevar DESC;"
 
    Set qdf = db.CreateQueryDef("Evening_Gluclose_Reading", newSQL)
End Sub