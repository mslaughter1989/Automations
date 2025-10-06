Attribute VB_Name = "ValidateLastName"
Sub ValidateLastNameColumn_SaveToLogsFolder_KeepOpen()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastNameCol As Long
    Dim cell As Range
    Dim resultWB As Workbook
    Dim resultWS As Worksheet
    Dim logRow As Long
    Dim currentValue As String
    Dim logFolder As String
    Dim logFileName As String
    Dim fullPath As String

    ' === Set your destination folder ===
    logFolder = "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Customer Success\Client Deployment\Kickoff Call\Automation\Logs\"
    
    ' Validate that the folder exists
    If Dir(logFolder, vbDirectory) = "" Then
        MsgBox "Folder not found: " & logFolder, vbCritical
        Exit Sub
    End If

    ' Use the active sheet (your opened CSV)
    Set ws = ActiveSheet

    ' Find the "Last Name" header column in row 1
    On Error Resume Next
    lastNameCol = Application.WorksheetFunction.match("Last Name", ws.Rows(1), 0)
    On Error GoTo 0

    If lastNameCol = 0 Then
        MsgBox "'Last Name' header not found in row 1.", vbExclamation
        Exit Sub
    End If

    ' Determine last row in the Last Name column
    lastRow = ws.Cells(ws.Rows.count, lastNameCol).End(xlUp).Row

    ' Create new workbook for logging
    Set resultWB = Workbooks.Add(xlWBATWorksheet)
    Set resultWS = resultWB.Sheets(1)
    resultWS.Name = "Validation Log"
    resultWS.Range("A1:E1").Value = Array("Row", "Column", "Cell Value", "Check Type", "Result")
    logRow = 2

    ' Loop through each cell in the Last Name column
    For Each cell In ws.Range(ws.Cells(2, lastNameCol), ws.Cells(lastRow, lastNameCol))
        currentValue = Trim(cell.Value)

        ' Check for blank
        If currentValue = "" Then
            resultWS.Cells(logRow, 1).Value = cell.Row
            resultWS.Cells(logRow, 2).Value = cell.Column
            resultWS.Cells(logRow, 3).Value = "(Blank)"
            resultWS.Cells(logRow, 4).Value = "Blank Check"
            resultWS.Cells(logRow, 5).Value = "Failed"
            logRow = logRow + 1
        End If

        ' Check for non-alphanumeric characters
        If currentValue <> "" Then
            If currentValue Like "*[!A-Za-z0-9]*" Then
                resultWS.Cells(logRow, 1).Value = cell.Row
                resultWS.Cells(logRow, 2).Value = cell.Column
                resultWS.Cells(logRow, 3).Value = currentValue
                resultWS.Cells(logRow, 4).Value = "Alphanumeric Check"
                resultWS.Cells(logRow, 5).Value = "Failed"
                logRow = logRow + 1
            End If
        End If

        ' Check for length over 50 characters
        If Len(currentValue) > 50 Then
            resultWS.Cells(logRow, 1).Value = cell.Row
            resultWS.Cells(logRow, 2).Value = cell.Column
            resultWS.Cells(logRow, 3).Value = currentValue
            resultWS.Cells(logRow, 4).Value = "Length Check"
            resultWS.Cells(logRow, 5).Value = "Failed"
            logRow = logRow + 1
        End If
    Next cell

    ' Save the result log as CSV in specified folder
    logFileName = "ValidationLog_" & Format(Now, "yyyymmdd_HHmmss") & ".csv"
    fullPath = logFolder & logFileName

    Application.DisplayAlerts = False
    resultWB.SaveAs fileName:=fullPath, fileFormat:=xlCSV
    Application.DisplayAlerts = True

    MsgBox "Validation complete." & vbCrLf & "Log saved to:" & vbCrLf & fullPath, vbInformation

    ' Keep the log workbook open
    resultWB.Activate

End Sub

