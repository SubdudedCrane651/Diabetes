'For Access
Option Compare Database

Public Sub AddDataToExcel(datestart As Integer)
    Dim oExcel As Excel.Application
    Dim oExcelWrkBk As Excel.Workbook
    Dim oExcelWrSht As Excel.Worksheet
    Dim dbs As Database
    Dim rs As Recordset
    Dim strSQL As String
    Set dbs = CurrentDb
    Dim qd As QueryDef
    Dim Reading As Double
    Dim x As Integer
    Dim xstr As String
    
    'datestart = 17
    
    'Start Excel
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If Err <> 0 Then
        Err.Clear
         On Error Goto Error_Handler
        Set oExcel = CreateObject("Excel.Application")
    Else
         On Error Goto Error_Handler
    End If
    
    Dim XL As Object, wb As Object
    Set XL = CreateObject("Excel.Application")
    With XL
        .Visible = False
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Workbooks.Open("C:\Users\rchrd\Documents\Richard\Diabetes.xlsm")
        .Run "DeleteSelection" 'Code stops here!
        .ActiveWorkbook.Save
        .ActiveWorkbook.Close(True)
        '.ActiveWorkbook.Quite
    End With
    Set XL = Nothing
    
    oExcel.ScreenUpdating = False
    oExcel.Visible = False 'This is false by default anyway
    
    Set oExcelWrkBk = oExcel.Workbooks.Open("C:\Users\rchrd\Documents\Richard\Diabetes.xlsm")
    Set oExcelWrSht = oExcelWrkBk.Sheets(1)
    
    
    oExcelWrkBk.RunAutoMacros(xlAutoActivate)
    
    x = 5
    
    oExcelWrSht.Range("n2").Value = datestart
    
    'Mourning Reading
    oExcelWrSht.Range("a3").Value = "Richard's Mourning Glucose Reading"
    Set qd = dbs.QueryDefs("Mourning_Gluclose_Reading")
    Set rst = qd.OpenRecordset
    
    Debug.Print(qd.SQL)
    
    ' Loop through records to calculate daily averages
    Do Until rst.EOF
        currDate = rst!Datevar
        
        ' If we're moving to a new date, store the previous day's average
        If currDate <> PrevDate And PrevDate <> 0 Then
            ' Insert averaged readings into Excel (only one row per day)
            oExcelWrSht.Cells(x, 1).Value = Format(PrevDate, "MM/DD/YYYY")
            oExcelWrSht.Cells(x, 2).Value = "02:00:00 AM" ' Fixed time
            oExcelWrSht.Cells(x, 3).Value = Round(sumReading / countReading, 1) ' Average reading rounded to 1 decimal place

            x = x + 1 ' Move to the next row
            
            ' Reset daily values
            sumReading = 0
            countReading = 0
        End If
        
        ' Add current reading to sum
        sumReading = sumReading + rst!Reading
        countReading = countReading + 1
        
        ' Update previous date tracker
        PrevDate = currDate
        
        rst.MoveNext
    Loop

    oExcelWrSht.Cells(x, 1).Value = Format(PrevDate, "MM/DD/YYYY")
    oExcelWrSht.Cells(x, 2).Value = "02:00:00 AM" ' Fixed time
    oExcelWrSht.Cells(x, 3).Value = Round(sumReading / countReading, 1) ' Average reading rounded to 1 decimal place

    rst.Close
    
    x = 5
    
    'Afternoon Reading
    oExcelWrSht.Range("e3").Value = "Richard's Afternoon Glucose Reading"
    Set qd = dbs.QueryDefs("Afternoon_Glucose_Reading")
    Set rst = qd.OpenRecordset
    
    Debug.Print(qd.SQL)
    
    PrevDate = Empty
    Doit = True
    
    ' Loop through records to calculate daily averages
    Do Until rst.EOF
        currDate = rst!Datevar
        currTime = rst!Timevar

        ' If we're moving to a new date, store the previous day's average
        If currDate <> PrevDate And PrevDate <> 0 Then
            ' Insert averaged readings into Excel (only one row per day)
            oExcelWrSht.Cells(x, 5).Value = Format(PrevDate, "MM/DD/YYYY")
            oExcelWrSht.Cells(x, 6).Value = "2:00 PM" ' Fixed time for afternoon
            ' Correct time insertion logic
            oExcelWrSht.Cells(x, 7).Value = Round(sumReading / countReading, 1) ' Average reading rounded to 1 decimal place
            x = x + 1
            oExcelWrSht.Cells(x, 5).Value = Format(PrevDate, "MM/DD/YYYY")
            oExcelWrSht.Cells(x, 6).Value = "12:00 PM" ' Fixed time for before lunch
            ' Correct time insertion logic
            If lunchcount > 0 Then
                oExcelWrSht.Cells(x, 7).Value = Round(lunchreading / lunchcount, 1) ' Average reading rounded to 1 decimal place
            End If

            x = x + 1 ' Move to the next row
            
            ' Reset daily values
            sumReading = 0
            countReading = 0
            lunchreading = 0
            lunchcount = 0
        End If
        
        If TimeValue(currTime) >= TimeValue("11:00:00 AM") And TimeValue(currTime) < TimeValue("1:00:00 PM") Then
            lunchreading = lunchreading + rst!Reading
            lunchcount = lunchcount + 1
        Else
            ' Add current reading to sum
            sumReading = sumReading + rst!Reading
            countReading = countReading + 1
        End If
        
        ' Update previous date tracker
        PrevDate = currDate
        
        rst.MoveNext

    Loop

    oExcelWrSht.Cells(x, 5).Value = Format(PrevDate, "MM/DD/YYYY")
    oExcelWrSht.Cells(x, 6).Value = "2:00 PM" ' Fixed time for afternoon
    ' Correct time insertion logic
    oExcelWrSht.Cells(x, 7).Value = Round(sumReading / countReading, 1) ' Average reading rounded to 1 decimal place
    x = x + 1
    oExcelWrSht.Cells(x, 5).Value = Format(PrevDate, "MM/DD/YYYY")
    oExcelWrSht.Cells(x, 6).Value = "12:00 PM" ' Fixed time for before lunch
    ' Correct time insertion logic
    If lunchcount > 0 Then
        oExcelWrSht.Cells(x, 7).Value = Round(lunchreading / lunchcount, 1) ' Average reading rounded to 1 decimal place
    End If

    rst.Close

    
    x = 5
    
    'Evening Reading
    oExcelWrSht.Range("i3").Value = "Richard's Evening Glucose Reading"
    Set qd = dbs.QueryDefs("Evening_Gluclose_Reading")
    Set rst = qd.OpenRecordset
    
    Debug.Print(qd.SQL)
    
    PrevDate = Empty
    
    ' Loop through records to calculate daily averages
    Do Until rst.EOF
        currDate = rst!Datevar
        
        ' If we're moving to a new date, store the previous day's average
        If currDate <> PrevDate And PrevDate <> 0 Then
            ' Insert averaged readings into Excel (only one row per day)
            oExcelWrSht.Cells(x, 9).Value = Format(PrevDate, "MM/DD/YYYY")
            oExcelWrSht.Cells(x, 10).Value = "10:00:00 PM" ' Fixed time
            oExcelWrSht.Cells(x, 11).Value = Round(sumReading / countReading, 1) ' Average reading rounded to 1 decimal place

            x = x + 1 ' Move to the next row
            
            ' Reset daily values
            sumReading = 0
            countReading = 0
        End If
        
        ' Add current reading to sum
        sumReading = sumReading + rst!Reading
        countReading = countReading + 1
        
        ' Update previous date tracker
        PrevDate = currDate
        
        rst.MoveNext
    Loop

    oExcelWrSht.Cells(x, 9).Value = Format(PrevDate, "MM/DD/YYYY")
    oExcelWrSht.Cells(x, 10).Value = "10:00:00 PM" ' Fixed time
    oExcelWrSht.Cells(x, 11).Value = Round(sumReading / countReading, 1) ' Average reading rounded to 1 decimal place
    
    rst.Close
    
    oExcelWrSht.Range("A1").Select
    
    oExcelWrkBk.Save
    
    oExcel.ScreenUpdating = True
    oExcel.Visible = True
    
    Exit_Point :
    Set oExcelWrSht = Nothing
    Set oExcelWrkBk = Nothing
    Set oExcel = Nothing
    Set rst = Nothing
    Exit Sub
    
    Error_Handler :
    MsgBox Err & " - " & Err.Description
    Goto Exit_Point
End Sub

Function ConvertFrenchDate(originalDate As String) As String
    Dim dateParts As Variant
    Dim formattedDateTime As String
    Dim sep As String

    ' Determine separator used ("/" or "-")
    If InStr(originalDate, "/") > 0 Then
        sep = "/"
    ElseIf InStr(originalDate, "-") > 0 Then
        sep = "-"
    Else
        MsgBox "Invalid date format!", vbCritical, "Error"
        Exit Function
    End If

    ' Split date and time separately
    dateParts = Split(originalDate, " ")

    ' Extract date components based on separator
    Dim dateValues As Variant
    dateValues = Split(dateParts(0), sep)

    ' Convert DD/MM/YYYY or DD-MM-YYYY into MM/DD/YYYY format
    formattedDateTime = Format(DateSerial(CInt(dateValues(2)), CInt(dateValues(1)), CInt(dateValues(0))), "MM/DD/YYYY") & " " & dateParts(1)

    ' Return corrected date and time
    ConvertFrenchDate = formattedDateTime
End Function

Sub ImportCSVToAccess()
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim csvFile As String
    Dim csvLine As String
    Dim dataArray As Variant
    Dim fileNum As Integer
    Dim lineCounter As Integer

    ' Define CSV file path
    csvFile = "D:\GlucoseReadings.csv" ' Ensure correct path

    ' Open the Access database
    Set db = CurrentDb()

    ' **Clear the table before inserting new data**
    db.Execute "DELETE FROM GlucoseReadings;", dbFailOnError
    MsgBox "Table cleared successfully!", vbInformation, "Reset Complete"
    
    ' Open a recordset to insert data
    Set rst = db.OpenRecordset("GlucoseReadings", dbOpenDynaset)
    
    ' Open the CSV file for reading
    fileNum = FreeFile
    Open csvFile For Input As #fileNum

        ' **Skip the first 2 lines**
        For lineCounter = 1 To 2
            Line Input #fileNum, csvLine ' Read and discard first 2 lines
            Debug.Print "Skipping line: " & csvLine ' Debugging check
        Next lineCounter
        
        ' Loop through CSV and insert records
        Do Until EOF(fileNum)
             On Error Goto HandleError ' Enable error handling

            Line Input #fileNum, csvLine
            dataArray = Split(csvLine, ",") ' Split CSV into array

            Dim originalDate As String
            Dim formattedDate As String
            
            originalDate = dataArray(2) ' French format (DD/MM/YYYY HH:MM or DD-MM-YYYY HH:MM)
            formattedDate = ConvertFrenchDate(originalDate) ' Convert to MM/DD/YYYY HH:MM
            
            rst.AddNew
            rst!Date = CDate(formattedDate) ' Store corrected date with time
            rst!GlucoseLevel = Format(CDbl(dataArray(4)), "0.0") ' Convert to Double (1 decimal)
            rst.Update
        Loop

        HandleError :
        ' Close connections (if no error)
        Close #fileNum
        rst.Close
        db.Close

        MsgBox "Data successfully imported!", vbInformation, "Import Complete"
        Call DeleteDiabetesFromFirstDate
        Call InsertGlucoseIntoDiabetes
    End Sub
    
    Sub InsertGlucoseIntoDiabetes()
        Dim db As DAO.Database
        Dim rstSource As DAO.Recordset
        Dim rstTarget As DAO.Recordset
        Dim glucoseDate As Date
        Dim glucoseTime As String
        Dim glucoseReading As Double

        ' Open the Access database
        Set db = CurrentDb()

        ' Open the source table (GlucoseReadings)
        Set rstSource = db.OpenRecordset("GlucoseReadings", dbOpenDynaset)

        ' Open the target table (Diabetes)
        Set rstTarget = db.OpenRecordset("Diabetes", dbOpenDynaset)

        ' Loop through GlucoseReadings table and insert into Diabetes table
        Do While Not rstSource.EOF
            ' Extract Date and Time separately
            glucoseDate = rstSource !Date
            glucoseTime = Format(rstSource !Date, "HH:MM:SS AM/PM") ' Extract time
            glucoseReading = rstSource !GlucoseLevel

            ' Insert data into Diabetes table
            rstTarget.AddNew
            rstTarget !Datevar = glucoseDate
            rstTarget !Timevar = glucoseTime
            rstTarget !Reading = Format(glucoseReading, "0.0") ' One decimal place
            rstTarget.Update

            ' Move to the next record
            rstSource.MoveNext
        Loop

        ' Close connections
        rstSource.Close
        rstTarget.Close
        db.Close

        MsgBox "Glucose readings successfully inserted into Diabetes table!", vbInformation, "Success"
    End Sub
    
    Function GetFirstDateFromCSV(csvPath As String) As String
        Dim fileNum As Integer
        Dim csvLine As String
        Dim firstDate As String
        Dim dataArray As Variant
        
        fileNum = FreeFile
        Open csvPath For Input As #fileNum

            ' Skip the first 2 header lines
            Line Input #fileNum, csvLine
            Line Input #fileNum, csvLine

            ' Get the **first actual data date**
            Line Input #fileNum, csvLine
            dataArray = Split(csvLine, ",")
            firstDate = dataArray(2) ' Extracts first date
            
            Close #fileNum

            ' Return the first date
            GetFirstDateFromCSV = firstDate
        End Function
        
        Sub DeleteDiabetesFromFirstDate()
            Dim db As DAO.Database
            Dim firstDate As String

            ' Get the first date from CSV
            firstDate = GetFirstDateFromCSV("D:\GlucoseReadings.csv")

            ' Open Access database
            Set db = CurrentDb()

            ' **Delete all records from first CSV date onward**
            db.Execute "DELETE FROM Diabetes WHERE Datevar >= #" & firstDate & "#;", dbFailOnError

            MsgBox "Deleted records from " & firstDate & " onwards in Diabetes table!", vbInformation, "Records Removed"

            db.Close
        End Sub