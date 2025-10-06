Attribute VB_Name = "ValidateFN_LN"
Sub ValidateFirstAndLastName()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colIndex As Long
    Dim cell As Range
    Dim resultWB As Workbook
    Dim resultWS As Worksheet
    Dim logRow As Long
    Dim currentValue As String
    Dim logFolder As String
    Dim logFileName As String
    Dim fullPath As String
    Dim headersToCheck As Variant
    Dim maxLengths As Variant
    Dim i As Integer
    Dim colLabel As String

    ' === Folder path to save log ===
    logFolder = "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Customer Success\Client Deployment\Kickoff Call\Automation\Logs\"
    If Right(logFolder, 1) <> "\" Then logFolder = logFolder & "\"

    If Dir(logFolder, vbDirectory) = "" Then
        MsgBox "Folder not found: " & logFolder, vbCritical
        Exit Sub
    End If

    ' Columns and max lengths
    headersToCheck = Array("First Name", "Last Name")
    maxLengths = Array(50, 50)

    Set ws = ActiveSheet

    ' Create log workbook
    Set resultWB = Workbooks.Add(xlWBATWorksheet)
    Set resultWS = resultWB.Sheets(1)
    resultWS.name = "Validation Log"
    resultWS.Range("A1:E1").Value = Array("Row", "Column", "Cell Value", "Check Type", "Result")
    logRow = 2

    ' Loop over "First Name" and "Last Name"
    For i = LBound(headersToCheck) To UBound(headersToCheck)
        colLabel = headersToCheck(i)

        On Error Resume Next
        colIndex = Application.WorksheetFunction.match(colLabel, ws.Rows(1), 0)
        On Error GoTo 0

        If colIndex = 0 Then
            MsgBox "'" & colLabel & "' header not found.", vbExclamation
            GoTo NextColumn
        End If

        lastRow = ws.Cells(ws.Rows.count, colIndex).End(xlUp).Row

        For Each cell In ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex))
            currentValue = cell.Value

            ' === Blank Check ===
            If Trim(currentValue) = "" Then
                LogIssue resultWS, logRow, cell.Row, colLabel, "(Blank)", "Blank Check", "Failed"
                logRow = logRow + 1
                GoTo NextCell
            End If

            ' === Alphanumeric w/ hyphen, apostrophe, space (no leading/trailing spaces) ===
            If Not IsValidNameFormat(currentValue) Then
                LogIssue resultWS, logRow, cell.Row, colLabel, currentValue, "Alphanumeric Check", "Failed"
                logRow = logRow + 1
            End If

            ' === Length Check ===
            If Len(Trim(currentValue)) > maxLengths(i) Then
                LogIssue resultWS, logRow, cell.Row, colLabel, currentValue, "Length Check", "Failed"
                logRow = logRow + 1
            End If

NextCell:
        Next cell

NextColumn:
    Next i

    ' === Save log CSV ===
    logFileName = "ValidationLog_FirstLastName_" & Format(Now, "yyyymmdd_HHmmss") & ".csv"
    fullPath = logFolder & logFileName

    Application.DisplayAlerts = False
    resultWB.SaveAs fileName:=fullPath, fileFormat:=xlCSV
    Application.DisplayAlerts = True

    MsgBox "Validation complete." & vbCrLf & "Log saved to:" & vbCrLf & fullPath, vbInformation
    resultWB.Activate

End Sub

' === Logging Helper ===
Sub LogIssue(ws As Worksheet, rowNum As Long, dataRow As Long, colName As String, val As String, checkType As String, result As String)
    ws.Cells(rowNum, 1).Value = dataRow
    ws.Cells(rowNum, 2).Value = colName
    ws.Cells(rowNum, 3).Value = val
    ws.Cells(rowNum, 4).Value = checkType
    ws.Cells(rowNum, 5).Value = result
End Sub

' === Name Format Helper (alphanumeric + hyphen + space + apostrophe, no leading/trailing spaces) ===
Function IsValidNameFormat(name As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .pattern = "^[A-Za-z0-9]+([ '-][A-Za-z0-9]+)*$"
        .IgnoreCase = True
        .Global = False
    End With

    IsValidNameFormat = regex.Test(name)
End Function

