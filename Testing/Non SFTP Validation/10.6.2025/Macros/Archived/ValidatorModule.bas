
Attribute VB_Name = "ValidatorModule"

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
        Set wb = Workbooks.Open(folderPath & csvFile)
        DoEvents
        baseName = Left(csvFile, InStrRev(csvFile, ".") - 1)
        Call Validate_WithRequiredColumnsSplitLog(baseName)
        wb.Close SaveChanges:=False
        DoEvents
        csvFile = Dir
    Loop

    MsgBox "Validation run complete for all CSV files.", vbInformation
End Sub

' ==== (Paste full Validate_WithRequiredColumnsSplitLog macro below) ====


<<<MARKER_FOR_FULL_MACRO_REPLACEMENT>>>