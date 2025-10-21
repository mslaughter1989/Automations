Attribute VB_Name = "DiagnosticModule_FIXED"
Option Explicit

' ===================================================================
' DIAGNOSTIC MODULE - FIXED VERSION
' Run this to find the exact error location
' ===================================================================

Public Sub DiagnoseValidationError()
    Dim sFilePath As String
    Dim sFileName As String
    Dim oFileInfo As FileInfo
    Dim oMapping As ColumnMapping
    Dim colRules As Collection
    Dim vCSVData As Variant
    Dim oResult As ValidationResult
    
    On Error GoTo ErrorHandler
    
    ' Step 1: Get a test file
    Debug.Print "=== DIAGNOSTIC START ==="
    Debug.Print "Step 1: Selecting file..."
    
    With Application.fileDialog(msoFileDialogFilePicker)
        .Title = "Select ONE CSV File for Diagnosis"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            sFilePath = .SelectedItems(1)
        Else
            Debug.Print "No file selected. Exiting."
            Exit Sub
        End If
    End With
    
    Debug.Print "File selected: " & sFilePath
    
    ' Step 2: Extract filename
    Debug.Print vbCrLf & "Step 2: Extracting filename..."
    sFileName = GetFileNameFromPath(sFilePath)
    Debug.Print "Filename: " & sFileName
    
    ' Step 3: Match pattern
    Debug.Print vbCrLf & "Step 3: Matching filename pattern..."
    oFileInfo = MatchFilenamePattern(sFileName)
    
    If Not oFileInfo.isValid Then
        Debug.Print "ERROR: No pattern match found!"
        Exit Sub
    End If
    
    Debug.Print "Pattern matched!"
    Debug.Print "  FileType: " & oFileInfo.fileType
    Debug.Print "  GroupID: " & oFileInfo.groupID
    
    ' Step 4: Get column mapping
    Debug.Print vbCrLf & "Step 4: Loading column mapping..."
    oMapping = GetColumnMapping(oFileInfo.fileType)
    
    If oMapping.fileType = "" Then
        Debug.Print "ERROR: No column mapping found for FileType: " & oFileInfo.fileType
        Exit Sub
    End If
    
    Debug.Print "Column mapping loaded!"
    Debug.Print "  FirstName column: " & oMapping.FirstName
    Debug.Print "  LastName column: " & oMapping.LastName
    Debug.Print "  MemberID column: " & oMapping.memberID
    Debug.Print "  GroupID column: " & oMapping.groupID
    Debug.Print "  ServiceOffering column: " & oMapping.serviceOffering
    Debug.Print "  EffectiveEndDate column: " & oMapping.effectiveEndDate
    Debug.Print "  Address2 column: " & oMapping.Address2
    
    ' Step 5: Load validation rules
    Debug.Print vbCrLf & "Step 5: Loading validation rules..."
    Set colRules = LoadValidationRules()
    
    If colRules Is Nothing Then
        Debug.Print "ERROR: Failed to load validation rules!"
        Exit Sub
    End If
    
    Debug.Print "Validation rules loaded: " & colRules.count & " rules"
    
    ' Step 6: Read CSV
    Debug.Print vbCrLf & "Step 6: Reading CSV file..."
    vCSVData = ReadCSVToArray(sFilePath)
    
    If IsEmpty(vCSVData) Then
        Debug.Print "ERROR: Failed to read CSV or file is empty!"
        Exit Sub
    End If
    
    Debug.Print "CSV read successfully!"
    Debug.Print "  Total rows: " & UBound(vCSVData, 1)
    Debug.Print "  Total columns: " & UBound(vCSVData, 2)
    
    ' Step 7: Initialize ValidationResult
    Debug.Print vbCrLf & "Step 7: Initializing ValidationResult object..."
    Set oResult = New ValidationResult
    oResult.fileName = sFileName
    oResult.filePath = sFilePath
    oResult.fileType = oFileInfo.fileType
    oResult.groupID = oFileInfo.groupID
    oResult.ProcessedDate = Now
    Debug.Print "ValidationResult initialized successfully!"
    
    ' Step 8: Test ValidateRowFields on FIRST DATA ROW ONLY
    Debug.Print vbCrLf & "Step 8: Testing field validation on row 2 (first data row)..."
    
    On Error GoTo RowError
    
    ' Test each field individually
    Debug.Print "  Testing FirstName..."
    If oMapping.FirstName > 0 And oMapping.FirstName <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.FirstName), "FirstName", 2, colRules, oResult)
        Debug.Print "  FirstName OK"
    Else
        Debug.Print "  FirstName column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing LastName..."
    If oMapping.LastName > 0 And oMapping.LastName <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.LastName), "LastName", 2, colRules, oResult)
        Debug.Print "  LastName OK"
    Else
        Debug.Print "  LastName column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing DOB..."
    If oMapping.DOB > 0 And oMapping.DOB <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.DOB), "DOB", 2, colRules, oResult)
        Debug.Print "  DOB OK"
    Else
        Debug.Print "  DOB column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing Gender..."
    If oMapping.Gender > 0 And oMapping.Gender <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.Gender), "Gender", 2, colRules, oResult)
        Debug.Print "  Gender OK"
    Else
        Debug.Print "  Gender column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing ZipCode..."
    If oMapping.ZipCode > 0 And oMapping.ZipCode <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.ZipCode), "ZipCode", 2, colRules, oResult)
        Debug.Print "  ZipCode OK"
    Else
        Debug.Print "  ZipCode column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing Address1..."
    If oMapping.Address1 > 0 And oMapping.Address1 <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.Address1), "Address1", 2, colRules, oResult)
        Debug.Print "  Address1 OK"
    Else
        Debug.Print "  Address1 column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing Address2..."
    If oMapping.Address2 > 0 And oMapping.Address2 <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.Address2), "Address2", 2, colRules, oResult)
        Debug.Print "  Address2 OK"
    Else
        Debug.Print "  Address2 column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing City..."
    If oMapping.City > 0 And oMapping.City <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.City), "City", 2, colRules, oResult)
        Debug.Print "  City OK"
    Else
        Debug.Print "  City column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing State..."
    If oMapping.State > 0 And oMapping.State <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.State), "State", 2, colRules, oResult)
        Debug.Print "  State OK"
    Else
        Debug.Print "  State column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing EffectiveDate..."
    If oMapping.EffectiveDate > 0 And oMapping.EffectiveDate <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.EffectiveDate), "EffectiveDate", 2, colRules, oResult)
        Debug.Print "  EffectiveDate OK"
    Else
        Debug.Print "  EffectiveDate column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing ServiceOffering..."
    If oMapping.serviceOffering > 0 And oMapping.serviceOffering <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.serviceOffering), "ServiceOffering", 2, colRules, oResult)
        Debug.Print "  ServiceOffering OK"
    Else
        Debug.Print "  ServiceOffering column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing MemberID..."
    If oMapping.memberID > 0 And oMapping.memberID <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.memberID), "MemberID", 2, colRules, oResult)
        Debug.Print "  MemberID OK"
    Else
        Debug.Print "  MemberID column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing GroupID..."
    If oMapping.groupID > 0 And oMapping.groupID <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.groupID), "GroupID", 2, colRules, oResult)
        Debug.Print "  GroupID OK"
    Else
        Debug.Print "  GroupID column not mapped or out of bounds!"
    End If
    
    Debug.Print "  Testing EffectiveEndDate..."
    If oMapping.effectiveEndDate > 0 And oMapping.effectiveEndDate <= UBound(vCSVData, 2) Then
        Call VALIDATION_FIX_COMPLETE.ValidateField(vCSVData(2, oMapping.effectiveEndDate), "EffectiveEndDate", 2, colRules, oResult)
        Debug.Print "  EffectiveEndDate OK"
    Else
        Debug.Print "  EffectiveEndDate column not mapped or out of bounds!"
    End If
    
    Debug.Print vbCrLf & "Step 9: Testing ActiveChecker_FIXED..."
    
    If oMapping.effectiveEndDate > 0 And oMapping.effectiveEndDate <= UBound(vCSVData, 2) Then
        Dim sEndDate As String
        sEndDate = Trim(CStr(vCSVData(2, oMapping.effectiveEndDate)))
        Debug.Print "  EffectiveEndDate value: '" & sEndDate & "'"
        
        Dim bActive As Boolean
        bActive = VALIDATION_FIX_COMPLETE.ActiveChecker_FIXED(sEndDate)
        Debug.Print "  ActiveChecker result: " & bActive
    Else
        Debug.Print "  EffectiveEndDate column not mapped - testing with blank value"
        bActive = VALIDATION_FIX_COMPLETE.ActiveChecker_FIXED("")
        Debug.Print "  ActiveChecker result for blank: " & bActive
    End If
    
    ' Step 10: Test Dictionary creation
    Debug.Print vbCrLf & "Step 10: Testing Dictionary object creation..."
    Dim testDict As Object
    Set testDict = CreateObject("Scripting.Dictionary")
    Debug.Print "  Dictionary created successfully!"
    Debug.Print "  Dictionary TypeName: " & TypeName(testDict)
    
    ' Step 11: Test CheckForDuplicates_FIXED
    Debug.Print vbCrLf & "Step 11: Testing CheckForDuplicates_FIXED..."
    On Error GoTo DuplicateError
    Call VALIDATION_FIX_COMPLETE.CheckForDuplicates_FIXED(vCSVData, oMapping, oResult, 2, UBound(vCSVData, 1))
    Debug.Print "  CheckForDuplicates_FIXED completed successfully!"
    
    Debug.Print vbCrLf & "=== DIAGNOSTIC COMPLETE - NO ERRORS FOUND ==="
    Debug.Print "Total Errors Found: " & oResult.ErrorCount
    Debug.Print "Total Warnings Found: " & oResult.WarningCount
    Debug.Print vbCrLf & "All functions are working correctly!"
    Debug.Print "The issue may be in how StartFileValidation calls these functions."
    
    Exit Sub
    
DuplicateError:
    Debug.Print "ERROR in CheckForDuplicates_FIXED!"
    Debug.Print "  Error Number: " & Err.Number
    Debug.Print "  Error Description: " & Err.Description
    Debug.Print "  Error Source: " & Err.Source
    Debug.Print vbCrLf & "*** THIS IS LIKELY YOUR MAIN PROBLEM! ***"
    Exit Sub
    
RowError:
    Debug.Print "ERROR in field validation!"
    Debug.Print "  Error Number: " & Err.Number
    Debug.Print "  Error Description: " & Err.Description
    Debug.Print "  Error Source: " & Err.Source
    Debug.Print vbCrLf & "*** THIS IS LIKELY YOUR MAIN PROBLEM! ***"
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR at early stage!"
    Debug.Print "  Error Number: " & Err.Number
    Debug.Print "  Error Description: " & Err.Description
    Debug.Print "  Error Source: " & Err.Source
End Sub
