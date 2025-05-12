Option Compare Database

Public Sub CreateDiabetesQrys(days)
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim newSQL As String
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "Mourning_Gluclose_Reading"
    On Error GoTo 0

    newSQL = "SELECT Format(Diabetes.Datevar,'MM/DD/YYYY') as Datevar, Diabetes.Timevar, Diabetes.Reading FROM Diabetes WHERE (((Diabetes.Datevar) >= Date() - " + Str(days) + ") And ((Diabetes.Timevar) <= #6:00:00 AM#) And ((Year([Diabetes].[Datevar])) = 2025)) ORDER BY Diabetes.Datevar DESC;"
 
    Set qdf = db.CreateQueryDef("Mourning_Gluclose_Reading", newSQL)
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "Afternoon_Glucose_Reading"
    On Error GoTo 0

    newSQL = "SELECT Format(Diabetes.Datevar,'MM/DD/YYYY') as Datevar, Diabetes.Timevar, Diabetes.Reading FROM Diabetes WHERE (((Diabetes.Datevar) >= Date() - " + Str(days) + ") And ((Diabetes.Timevar) >=#6:00:00 AM# And (Diabetes.Timevar)<=#18:00:00 PM#) And ((Year([Diabetes].[Datevar])) = 2025)) ORDER BY Diabetes.Datevar DESC;"
 
    Set qdf = db.CreateQueryDef("Afternoon_Glucose_Reading", newSQL)
    
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "Evening_Gluclose_Reading"
    On Error GoTo 0

    newSQL = "SELECT Format(Diabetes.Datevar,'MM/DD/YYYY') as Datevar, Diabetes.Timevar, Diabetes.Reading FROM Diabetes WHERE (((Diabetes.Datevar) >= Date() - " + Str(days) + ") And ((Diabetes.Timevar) >= #18:00:00 PM#) And ((Year([Diabetes].[Datevar])) = 2025)) ORDER BY Diabetes.Datevar DESC;"
 
    Set qdf = db.CreateQueryDef("Evening_Gluclose_Reading", newSQL)
End Sub

Public Sub Test()
MsgBox "Number of Days"
End Sub