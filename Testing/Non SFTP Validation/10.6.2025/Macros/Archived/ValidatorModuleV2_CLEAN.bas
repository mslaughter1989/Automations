



Sub RunValidationOnAllCSVsInFolder()
    Dim fDialog As FileDialog
    Dim folderPath As String
    Dim csvFile As String
    Dim wb As Workbook
    Dim baseName As String

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select Folder Containing CSV Files"
        If .Show <> -1 Then
            MsgBox "No folder selected. Macro canceled.", vbExclamation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    csvFile = Dir(folderPath & "*.csv")

    Do While csvFile <> ""
        Debug.Print "Processing: " & csvFile

        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=folderPath & csvFile)
        If wb Is Nothing Then
            MsgBox "Could not open file: " & folderPath & csvFile, vbExclamation
            csvFile = Dir
            On Error GoTo 0
            GoTo SkipFile
        End If
        On Error GoTo 0

        baseName = Left(csvFile, InStrRev(csvFile, ".") - 1)
        Call Validate_WithRequiredColumnsSplitLog(baseName)

        wb.Close SaveChanges:=False
        DoEvents

SkipFile:
        csvFile = Dir
    Loop

    MsgBox "Validation run complete for all CSV files.", vbInformation
End Sub


Sub Validate_WithRequiredColumnsSplitLog(Optional baseFileName As String)
    If baseFileName = "" Then baseFileName = "UnnamedFile"

    Dim ws As Worksheet, resultWB As Workbook
    Dim logAll As Worksheet, logRequired As Worksheet
    Dim logRowAll As Long, logRowReq As Long
    Dim colIndex As Long, lastRow As Long
    Dim colCell As Range, cell As Range
    Dim header As String, currentValue As String
    Dim logFolder As String, logFileName As String, fullPath As String
    Dim checkDict As Object, checkList As Variant
    Dim requiredCols As Object
    Dim i As Long

    logFolder = "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Customer Success\Client Deployment\Kickoff Call\Automation\Logs\"
    If Right(logFolder, 1) <> "\" Then logFolder = logFolder & "\"
    If Dir(logFolder, vbDirectory) = "" Then
        MsgBox "Folder not found: " & logFolder, vbCritical
        Exit Sub
    End If

    Set requiredCols = CreateObject("Scripting.Dictionary")
    Dim requiredList As Variant
    requiredList = Array("First Name", "Last Name", "Date of Birth", "E-mail Address", "Effective Start", _
                         "Member Type", "Client Member ID", "Client Primary Member ID", "Service Offering", _
                         "Group ID", "Group Name")
    For i = LBound(requiredList) To UBound(requiredList)
        requiredCols.Add requiredList(i), True
    Next i

    Set checkDict = CreateObject("Scripting.Dictionary")
    checkDict.Add "First Name", Array("Blank Check", "Name Format", 50)
    checkDict.Add "Last Name", Array("Blank Check", "Name Format", 50)
    checkDict.Add "Gender", Array("Blank Check", "M/F Only", 6)
    checkDict.Add "Date of Birth", Array("Blank Check", "Date Format", 10)
    checkDict.Add "Address Line 1", Array("Blank Check", "Address Format", 150)
    checkDict.Add "Address Line 2", Array(150)
    checkDict.Add "City", Array("Blank Check", "Name Format", 150)
    checkDict.Add "State", Array("Blank Check", 2)
    checkDict.Add "Zip Code", Array("Blank Check", "Zip Format", 10)
    checkDict.Add "Country Code", Array("Blank Check", 2)
    checkDict.Add "Mobile Phone", Array("Blank Check", "Phone Format")
    checkDict.Add "E-mail Address", Array("Blank Check", "Email Format", 150)
    checkDict.Add "Effective Start", Array("Blank Check", "Date Format", 10)
    checkDict.Add "Effective End", Array("Blank Check", "Date Format", 10)
    checkDict.Add "Member Type", Array("Blank Check", "Alpha Only", "ValidType", 7)
    checkDict.Add "Client Member ID", Array("Blank Check", "MinLen:6", 15)
    checkDict.Add "Secondary Client Member ID", Array("Blank Check", 50)
    checkDict.Add "Client Primary Member ID", Array("MaxLen:50")
    checkDict.Add "Service Offering", Array("Blank Check", 150)
    checkDict.Add "Group ID", Array("Blank Check", "MinLen:4", 50)
    checkDict.Add "Group Name", Array("Blank Check", "MinLen:4", 50)
    checkDict.Add "Meta Tag 1", Array("Blank Check", 150)
    checkDict.Add "Meta Tag 2", Array("Blank Check", 150)
    checkDict.Add "Meta Tag 3", Array("Blank Check", 150)
    checkDict.Add "Meta Tag 4", Array("Blank Check", 150)
    checkDict.Add "Meta Tag 5", Array("Blank Check", 150)

    Set ws = ActiveSheet
    Set resultWB = Workbooks.Add
    Set logRequired = resultWB.Sheets(1)
    logRequired.Name = "Required Fields Log"
    Set logAll = resultWB.Sheets.Add(After:=logRequired)
    logAll.Name = "All Validations Log"
    logRowAll = 2
    logRowReq = 2

    logAll.Range("A1:E1").Value = Array("Row", "Column", "Cell Value", "Check Type", "Result")
    logRequired.Range("A1:E1").Value = Array("Row", "Column", "Cell Value", "Check Type", "Result")

    For Each colCell In ws.Range("1:1")
        header = Trim(colCell.Value)
        If header = "" Then Exit For
        If checkDict.exists(header) Then
            colIndex = colCell.Column
            lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
            checkList = checkDict(header)

            For Each cell In ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex))
                currentValue = Trim(cell.Value)

                If IsInArray("Blank Check", checkList) Then
                    If currentValue = "" Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, "(Blank)", "Blank Check", requiredCols.exists(header)
                        GoTo NextCell
                    End If
                End If

                If IsInArray("Name Format", checkList) Then
                    If Not RegexTest(currentValue, "^[A-Za-z0-9]+([ '\-][A-Za-z0-9]+)*$") Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Invalid Name Format", requiredCols.exists(header)
                    End If
                End If

                If IsInArray("Address Format", checkList) Then
                    If Not RegexTest(currentValue, "^[A-Za-z0-9]+([ '\-\.][A-Za-z0-9]+)*$") Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Invalid Address Format", requiredCols.exists(header)
                    End If
                End If

                If IsInArray("Email Format", checkList) Then
                    If Not RegexTest(currentValue, "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$") Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Invalid Email Format", requiredCols.exists(header)
                    End If
                End If

                If IsInArray("Date Format", checkList) Then
                    If Not RegexTest(currentValue, "^((0?[1-9]|1[0-2])[/-](0?[1-9]|[12][0-9]|3[01])[/-](\d{4})|(\d{4})-(0[1-9]|1[0-2])-(0[1-9]|[12][0-9]|3[01]))$") Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Invalid Date Format", requiredCols.exists(header)
                    End If
                End If

                If IsInArray("Zip Format", checkList) Then
                    If Not RegexTest(currentValue, "^\d{5}(-\d{4})?$") Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Invalid Zip Format", requiredCols.exists(header)
                    End If
                End If

                If IsInArray("Phone Format", checkList) Then
                    If Not RegexTest(currentValue, "^\+?\d{0,2}[- ]?\(?\d{3}\)?[- ]?\d{3}[- ]?\d{4}$") Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Invalid Phone Format", requiredCols.exists(header)
                    End If
                End If

                If IsInArray("M/F Only", checkList) Then
                    If UCase(currentValue) <> "M" And UCase(currentValue) <> "F" Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "M/F Only Check", requiredCols.exists(header)
                    End If
                End If

                If IsInArray("Alpha Only", checkList) Then
                    If Not RegexTest(currentValue, "^[A-Za-z]+$") Then
                        LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Alpha Only", requiredCols.exists(header)
                    End If
                End If

                If IsInArray("ValidType", checkList) Then
                    Select Case UCase(currentValue)
                        Case "P", "PRIMARY", "S", "SPOUSE", "C", "CHILD", "OTHER"
                        Case Else
                            LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Invalid Member Type", requiredCols.exists(header)
                    End Select
                End If

                For Each item In checkList
                    If IsNumeric(item) Then
                        If Len(currentValue) > item Then
                            LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Max Length: " & item, requiredCols.exists(header)
                        End If
                    ElseIf Left(item, 7) = "MinLen:" Then
                        If Len(currentValue) < CLng(Mid(item, 8)) Then
                            LogToBoth logRequired, logAll, logRowReq, logRowAll, cell.Row, header, currentValue, "Min Length: " & Mid(item, 8), requiredCols.exists(header)
                        End If
                    End If
                Next item

NextCell:
            Next cell
        End If
    Next colCell

    logFileName = "ValidationLog_" & baseFileName & "_" & Format(Now, "yyyymmdd_HHmmss") & ".xlsx"
    fullPath = logFolder & logFileName
    Application.DisplayAlerts = False
    resultWB.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    resultWB.Activate
    MsgBox "Validation complete. Log saved to:" & vbCrLf & fullPath, vbInformation
End Sub

Function RegexTest(inputStr As String, pattern As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = pattern
        .IgnoreCase = True
        .Global = False
    End With
    RegexTest = regex.Test(inputStr)
End Function

Function IsInArray(val As Variant, arr As Variant) As Boolean
    Dim item As Variant
    For Each item In arr
        If item = val Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

Sub LogToBoth(wsReq As Worksheet, wsAll As Worksheet, ByRef rowReq As Long, ByRef rowAll As Long, _
              rowNum As Long, colName As String, val As String, checkType As String, isRequired As Boolean)

    wsAll.Cells(rowAll, 1).Value = rowNum
    wsAll.Cells(rowAll, 2).Value = colName
    wsAll.Cells(rowAll, 3).Value = val
    wsAll.Cells(rowAll, 4).Value = checkType
    wsAll.Cells(rowAll, 5).Value = "Failed"
    rowAll = rowAll + 1

    If isRequired Then
        wsReq.Cells(rowReq, 1).Value = rowNum
        wsReq.Cells(rowReq, 2).Value = colName
        wsReq.Cells(rowReq, 3).Value = val
        wsReq.Cells(rowReq, 4).Value = checkType
        wsReq.Cells(rowReq, 5).Value = "Failed"
        rowReq = rowReq + 1
    End If
End Sub
