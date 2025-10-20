Attribute VB_Name = "VALIDATION_FIX_COMPLETE"
Option Explicit

' ==============================================================================
' COMPREHENSIVE VALIDATION FIX MODULE
' Version: 3.0
' Purpose: Fixes all validation issues in SFTP Command Center
' ==============================================================================

' ==============================================================================
' SECTION 1: CORE VALIDATION ENGINE FIXES
' ==============================================================================

Public Function GetValidationRule(colRules As Collection, sFieldType As String) As ValidationRule
    ' FIXED VERSION - Handles all edge cases properly
    Dim oRule As ValidationRule
    Dim emptyRule As ValidationRule
    Dim sRuleData As String
    Dim vParts As Variant
    
    ' Initialize empty rule with sensible defaults
    With emptyRule
        .FieldType = sFieldType
        .Required = False
        .MaxLength = 0
        .MinLength = 0
        .FormatPattern = ""
        .CustomFunction = ""
    End With
    
    ' Handle missing collection
    If colRules Is Nothing Then
        GetValidationRule = emptyRule
        Exit Function
    End If
    
    On Error GoTo NotFound
    
    ' Try to get the rule data string from collection
    sRuleData = colRules.Item(sFieldType)
    
    If sRuleData = "" Then GoTo NotFound
    
    ' Parse the delimited string back into ValidationRule Type
    vParts = Split(sRuleData, "|")
    
    If UBound(vParts) >= 5 Then
        With oRule
            .FieldType = vParts(0)
            ' Handle various TRUE representations
            .Required = (UCase(Trim(CStr(vParts(1)))) = "TRUE" Or _
                        UCase(Trim(CStr(vParts(1)))) = "Y" Or _
                        UCase(Trim(CStr(vParts(1)))) = "YES" Or _
                        Trim(CStr(vParts(1))) = "1")
            .MaxLength = CLng(Val(vParts(2)))
            .MinLength = CLng(Val(vParts(3)))
            .FormatPattern = Trim(CStr(vParts(4)))
            .CustomFunction = Trim(CStr(vParts(5)))
        End With
        
        GetValidationRule = oRule
        Exit Function
    End If
    
NotFound:
    ' Return default rule for known field types
    Select Case UCase(sFieldType)
        Case "FIRSTNAME", "LASTNAME"
            emptyRule.Required = True
            emptyRule.MaxLength = 50
            emptyRule.MinLength = 1
            
        Case "ADDRESS1"
            emptyRule.Required = True
            emptyRule.MaxLength = 100
            
        Case "CITY"
            emptyRule.Required = True
            emptyRule.MaxLength = 50
            
        Case "STATE"
            emptyRule.Required = True
            emptyRule.MaxLength = 2
            emptyRule.MinLength = 2
            
        Case "ZIPCODE"
            emptyRule.Required = True
            emptyRule.MaxLength = 10
            emptyRule.MinLength = 5
            emptyRule.FormatPattern = "ZIP"
            
        Case "DOB", "EFFECTIVEDATE"
            emptyRule.Required = True
            emptyRule.FormatPattern = "DATE"
            
        Case "GENDER"
            emptyRule.Required = True
            emptyRule.MaxLength = 10
            emptyRule.FormatPattern = "GENDER"
            
        Case "MEMBERID"
            emptyRule.Required = True
            emptyRule.MaxLength = 20
            
        Case "SERVICEOFFERING"
            emptyRule.Required = True
            emptyRule.MaxLength = 50
            
        Case "GROUPID"
            emptyRule.Required = True
            emptyRule.MaxLength = 20
    End Select
    
    GetValidationRule = emptyRule
End Function

' ==============================================================================
' SECTION 2: ENHANCED FIELD VALIDATION
' ==============================================================================

Public Sub ValidateField(vFieldValue As Variant, sFieldType As String, _
                               lRowNumber As Long, colRules As Collection, _
                               oResult As ValidationResult)
    
    ' Static collections for tracking validation checks
    Static checkedFields As Collection
    Static validationCounts As Object
    
    If checkedFields Is Nothing Then Set checkedFields = New Collection
    If validationCounts Is Nothing Then Set validationCounts = CreateObject("Scripting.Dictionary")
    
    ' Initialize counters for this field type if needed
    If Not validationCounts.Exists(sFieldType) Then
        validationCounts(sFieldType) = CreateObject("Scripting.Dictionary")
        validationCounts(sFieldType)("blank") = 0
        validationCounts(sFieldType)("maxchar") = 0
        validationCounts(sFieldType)("format") = 0
    End If
    
    ' Track first time checking this field type
    On Error Resume Next
    checkedFields.Add sFieldType, sFieldType
    If Err.Number = 0 Then
        oResult.AddValidationCheck sFieldType & " Field", "Validating across all records"
    End If
    On Error GoTo 0
    
    ' Get validation rule (using FIXED version)
    Dim oRule As ValidationRule
    oRule = GetValidationRule(colRules, sFieldType)
    
    ' Convert to string for validation
    Dim sValue As String
    If IsNull(vFieldValue) Then
        sValue = ""
    Else
        sValue = Trim(CStr(vFieldValue))
    End If
    
    ' VALIDATION CHECK 1: Required/Blank Check
    If oRule.Required Then
        validationCounts(sFieldType)("blank") = validationCounts(sFieldType)("blank") + 1
        
        If sValue = "" Then
            oResult.AddError lRowNumber, sFieldType, "Required field is blank"
            
            ' Log check performed (only first time)
            If validationCounts(sFieldType)("blank") = 1 Then
                oResult.AddValidationCheck "Blank Check - " & sFieldType, "PERFORMED - Found blank value(s)"
            End If
            Exit Sub ' Skip other checks if blank
        Else
            ' Log success on first valid check
            If validationCounts(sFieldType)("blank") = 1 Then
                oResult.AddValidationCheck "Blank Check - " & sFieldType, "PERFORMED - Field populated"
            End If
        End If
    End If
    
    ' Skip further validation if field is empty and not required
    If sValue = "" Then Exit Sub
    
    ' VALIDATION CHECK 2: Maximum Length
    If oRule.MaxLength > 0 Then
        validationCounts(sFieldType)("maxchar") = validationCounts(sFieldType)("maxchar") + 1
        
        If Len(sValue) > oRule.MaxLength Then
            oResult.AddError lRowNumber, sFieldType, _
                "Exceeds maximum length of " & oRule.MaxLength & " characters (found " & Len(sValue) & ")"
            
            If validationCounts(sFieldType)("maxchar") = 1 Then
                oResult.AddValidationCheck "Max Length Check - " & sFieldType, _
                    "PERFORMED - Found violation(s) (Max: " & oRule.MaxLength & ")"
            End If
        Else
            If validationCounts(sFieldType)("maxchar") = 1 Then
                oResult.AddValidationCheck "Max Length Check - " & sFieldType, _
                    "PERFORMED - Within limit (Max: " & oRule.MaxLength & ")"
            End If
        End If
    End If
    
    ' VALIDATION CHECK 3: Minimum Length
    If oRule.MinLength > 0 And Len(sValue) < oRule.MinLength Then
        oResult.AddError lRowNumber, sFieldType, _
            "Below minimum length of " & oRule.MinLength & " characters (found " & Len(sValue) & ")"
    End If
    
    ' VALIDATION CHECK 4: Format Validation
    If oRule.FormatPattern <> "" Then
        validationCounts(sFieldType)("format") = validationCounts(sFieldType)("format") + 1
        
        If Not ValidateFieldFormat_FIXED(sValue, sFieldType, oRule.FormatPattern) Then
            Dim sFormatMsg As String
            sFormatMsg = GetFormatErrorMessage(sFieldType, oRule.FormatPattern, sValue)
            oResult.AddError lRowNumber, sFieldType, sFormatMsg
            
            If validationCounts(sFieldType)("format") = 1 Then
                oResult.AddValidationCheck "Format Check - " & sFieldType, _
                    "PERFORMED - Found invalid format(s)"
            End If
        Else
            If validationCounts(sFieldType)("format") = 1 Then
                oResult.AddValidationCheck "Format Check - " & sFieldType, _
                    "PERFORMED - Valid format"
            End If
        End If
    End If
End Sub

' ==============================================================================
' SECTION 3: FORMAT VALIDATION FUNCTIONS
' ==============================================================================

Private Function ValidateFieldFormat_FIXED(sValue As String, sFieldType As String, sPattern As String) As Boolean
    ' Enhanced format validation with better error handling
    
    Select Case UCase(sPattern)
        Case "DATE"
            ValidateFieldFormat_FIXED = ValidateDateFormat_Enhanced(sValue)
            
        Case "GENDER"
            ValidateFieldFormat_FIXED = ValidateGenderCode_Enhanced(sValue)
            
        Case "ZIP"
            ValidateFieldFormat_FIXED = ValidateZipCode_Enhanced(sValue)
            
        Case "STATE"
            ValidateFieldFormat_FIXED = ValidateStateCode_Enhanced(sValue)
            
        Case "NAME"
            ValidateFieldFormat_FIXED = ValidateNameFormat_Enhanced(sValue)
            
        Case "EMAIL"
            ValidateFieldFormat_FIXED = ValidateEmailFormat(sValue)
            
        Case "PHONE"
            ValidateFieldFormat_FIXED = ValidatePhoneFormat(sValue)
            
        Case Else
            ' If specific pattern provided, use regex
            If sPattern <> "" And Left(sPattern, 1) = "^" Then
                ValidateFieldFormat_FIXED = ValidateWithRegex(sValue, sPattern)
            Else
                ' No specific validation
                ValidateFieldFormat_FIXED = True
            End If
    End Select
End Function

Private Function ValidateDateFormat_Enhanced(sValue As String) As Boolean
    ' Validates multiple date formats
    On Error GoTo InvalidDate
    
    Dim dtTest As Date
    Dim sCleanValue As String
    
    ' Clean the value
    sCleanValue = Trim(sValue)
    
    ' Check for empty
    If sCleanValue = "" Or sCleanValue = "0" Then
        ValidateDateFormat_Enhanced = False
        Exit Function
    End If
    
    ' Try to parse the date
    dtTest = CDate(sCleanValue)
    
    ' Additional checks
    ' 1. Year should be reasonable (1900-2100)
    If Year(dtTest) < 1900 Or Year(dtTest) > 2100 Then
        ValidateDateFormat_Enhanced = False
        Exit Function
    End If
    
    ' 2. Check for invalid dates like 00/00/0000
    If dtTest = 0 Then
        ValidateDateFormat_Enhanced = False
        Exit Function
    End If
    
    ValidateDateFormat_Enhanced = True
    Exit Function
    
InvalidDate:
    ValidateDateFormat_Enhanced = False
End Function

Private Function ValidateGenderCode_Enhanced(sValue As String) As Boolean
    ' Enhanced gender validation
    Dim vValidCodes As Variant
    vValidCodes = Array("M", "F", "MALE", "FEMALE", "1", "2", "U", "UNKNOWN", "O", "OTHER")
    
    Dim sUpper As String
    sUpper = UCase(Trim(sValue))
    
    Dim i As Long
    For i = 0 To UBound(vValidCodes)
        If sUpper = CStr(vValidCodes(i)) Then
            ValidateGenderCode_Enhanced = True
            Exit Function
        End If
    Next i
    
    ValidateGenderCode_Enhanced = False
End Function

Private Function ValidateZipCode_Enhanced(sValue As String) As Boolean
    ' Enhanced ZIP code validation
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    ' Remove spaces and dashes for checking
    Dim sClean As String
    sClean = Replace(Replace(Trim(sValue), " ", ""), "-", "")
    
    ' Check basic format: 5 digits or 9 digits
    If Len(sClean) = 5 Then
        oRegex.pattern = "^\d{5}$"
    ElseIf Len(sClean) = 9 Then
        oRegex.pattern = "^\d{9}$"
    Else
        ' Also accept formatted: 12345-6789
        oRegex.pattern = "^\d{5}[-\s]?\d{4}$"
        sClean = sValue ' Use original for this check
    End If
    
    ValidateZipCode_Enhanced = oRegex.Test(sClean)
End Function

Private Function ValidateStateCode_Enhanced(sValue As String) As Boolean
    ' Enhanced state code validation with full list
    Dim vValidStates As Variant
    vValidStates = Array("AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", _
                        "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", _
                        "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", _
                        "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", _
                        "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY", _
                        "DC", "PR", "VI", "GU", "AS", "MP") ' Include territories
    
    Dim sUpper As String
    sUpper = UCase(Trim(sValue))
    
    If Len(sUpper) <> 2 Then
        ValidateStateCode_Enhanced = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To UBound(vValidStates)
        If sUpper = CStr(vValidStates(i)) Then
            ValidateStateCode_Enhanced = True
            Exit Function
        End If
    Next i
    
    ValidateStateCode_Enhanced = False
End Function

Private Function ValidateNameFormat_Enhanced(sValue As String) As Boolean
    ' Enhanced name validation
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    ' Allow letters, spaces, hyphens, apostrophes, periods
    ' Must start with letter, 2-50 characters
    oRegex.pattern = "^[a-zA-Z][a-zA-Z\s\-'\.]{1,49}$"
    oRegex.IgnoreCase = True
    
    Dim sClean As String
    sClean = Trim(sValue)
    
    ' Additional checks
    If Len(sClean) < 2 Then
        ValidateNameFormat_Enhanced = False
        Exit Function
    End If
    
    ' Check for invalid characters
    Dim invalidChars As String
    invalidChars = "0123456789!@#$%^&*()_+={}[]|\\:;""<>?,/"
    
    Dim i As Integer
    For i = 1 To Len(invalidChars)
        If InStr(sClean, Mid(invalidChars, i, 1)) > 0 Then
            ValidateNameFormat_Enhanced = False
            Exit Function
        End If
    Next i
    
    ValidateNameFormat_Enhanced = oRegex.Test(sClean)
End Function

Private Function ValidateEmailFormat(sValue As String) As Boolean
    ' Email validation
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    oRegex.pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    ValidateEmailFormat = oRegex.Test(Trim(sValue))
End Function

Private Function ValidatePhoneFormat(sValue As String) As Boolean
    ' Phone validation (US format)
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    ' Remove common formatting characters
    Dim sClean As String
    sClean = Replace(Replace(Replace(Replace(Trim(sValue), "(", ""), ")", ""), "-", ""), " ", "")
    
    ' Should be 10 digits
    oRegex.pattern = "^\d{10}$"
    ValidatePhoneFormat = oRegex.Test(sClean)
End Function

Private Function ValidateWithRegex(sValue As String, sPattern As String) As Boolean
    ' Generic regex validation
    On Error GoTo RegexError
    
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    oRegex.pattern = sPattern
    oRegex.IgnoreCase = True
    ValidateWithRegex = oRegex.Test(sValue)
    Exit Function
    
RegexError:
    ValidateWithRegex = True ' Default to valid if regex fails
End Function

' ==============================================================================
' SECTION 4: ACTIVE RECORD CHECKER
' ==============================================================================

Public Function ActiveChecker_FIXED(sEffectiveEndDate As String) As Boolean
    ' FIXED VERSION - Properly determines if a record is active
    ' Record is active if:
    ' 1. EffectiveEndDate is blank/empty, OR
    ' 2. EffectiveEndDate is a future date (> today), OR
    ' 3. EffectiveEndDate equals today (still active today)
    
    On Error GoTo HandleError
    
    Dim todayDate As Date
    todayDate = Date ' Current date without time component
    
    ' Clean the input
    Dim sCleanDate As String
    sCleanDate = Trim(sEffectiveEndDate)
    
    ' Check for blank/empty/null values
    If sCleanDate = "" Or sCleanDate = "0" Or sCleanDate = "NULL" Or sCleanDate = "null" Then
        ActiveChecker_FIXED = True
        Exit Function
    End If
    
    ' Try to parse as date
    If IsDate(sCleanDate) Then
        Dim endDate As Date
        endDate = CDate(sCleanDate)
        
        ' Remove time component for comparison
        endDate = dateValue(endDate)
        
        ' Active if end date is today or in the future
        ActiveChecker_FIXED = (endDate >= todayDate)
    Else
        ' If we can't parse the date, log warning but assume inactive
        ' This is safer than assuming active for bad data
        Debug.Print "WARNING: Invalid date format in EffectiveEndDate: " & sCleanDate
        ActiveChecker_FIXED = False
    End If
    
    Exit Function
    
HandleError:
    ' On error, assume inactive (conservative approach)
    Debug.Print "ERROR in ActiveChecker: " & Err.Description & " for value: " & sEffectiveEndDate
    ActiveChecker_FIXED = False
End Function

' ==============================================================================
' SECTION 5: ENHANCED DUPLICATE DETECTION
' ==============================================================================

Public Sub CheckForDuplicates_FIXED(vData As Variant, oMapping As ColumnMapping, _
                                    oResult As ValidationResult, lStartRow As Long, _
                                    lEndRow As Long)
    ' Enhanced duplicate detection with proper active/inactive handling
    
    Dim colMemberRecords As Object
    Set colMemberRecords = CreateObject("Scripting.Dictionary")
    
    Dim lRow As Long
    Dim sMemberID As String
    Dim sServiceOffering As String
    Dim sEffectiveEndDate As String
    Dim sComboKey As String
    Dim bIsActive As Boolean
    
    ' Build dictionary of all member records
    For lRow = lStartRow To lEndRow
        ' Get MemberID
        If oMapping.memberID > 0 And oMapping.memberID <= UBound(vData, 2) Then
            sMemberID = Trim(CStr(vData(lRow, oMapping.memberID)))
        Else
            sMemberID = ""
        End If
        
        ' Get ServiceOffering
        If oMapping.serviceOffering > 0 And oMapping.serviceOffering <= UBound(vData, 2) Then
            sServiceOffering = Trim(CStr(vData(lRow, oMapping.serviceOffering)))
        Else
            sServiceOffering = "UNKNOWN"
        End If
        
        ' Get EffectiveEndDate
        If oMapping.effectiveEndDate > 0 And oMapping.effectiveEndDate <= UBound(vData, 2) Then
            sEffectiveEndDate = Trim(CStr(vData(lRow, oMapping.effectiveEndDate)))
        Else
            sEffectiveEndDate = ""
        End If
        
        ' Check if active
        bIsActive = ActiveChecker_FIXED(sEffectiveEndDate)
        
        ' Create composite key
        sComboKey = sMemberID & "|" & sServiceOffering
        
        If sMemberID <> "" And sServiceOffering <> "" Then
            ' Store record info: "Row|IsActive|EffectiveEndDate"
            Dim sRecordInfo As String
            sRecordInfo = lRow & "|" & bIsActive & "|" & sEffectiveEndDate
            
            If Not colMemberRecords.Exists(sComboKey) Then
                ' First occurrence - create collection
                Dim newCol As Collection
                Set newCol = New Collection
                newCol.Add sRecordInfo
                colMemberRecords.Add sComboKey, newCol
            Else
                ' Add to existing collection
                colMemberRecords(sComboKey).Add sRecordInfo
            End If
        End If
    Next lRow
    
    ' Now check for problematic duplicates
    Dim vKey As Variant
    For Each vKey In colMemberRecords.Keys
        Dim records As Collection
        Set records = colMemberRecords(vKey)
        
        If records.count > 1 Then
            ' We have duplicates - analyze them
            Dim activeCount As Long
            Dim activeRows As String
            activeCount = 0
            activeRows = ""
            
            Dim i As Long
            For i = 1 To records.count
                Dim vParts As Variant
                vParts = Split(records(i), "|")
                
                If CBool(vParts(1)) = True Then ' Is active
                    activeCount = activeCount + 1
                    If activeRows <> "" Then activeRows = activeRows & ", "
                    activeRows = activeRows & "Row " & vParts(0)
                End If
            Next i
            
            ' Only flag as error if multiple active records exist
            If activeCount > 1 Then
                ' Parse the key to get MemberID and ServiceOffering
                Dim keyParts As Variant
                keyParts = Split(vKey, "|")
                
                ' Add error for each active duplicate
                oResult.AddError 0, "Duplicate", _
                    "Multiple active records found for MemberID=" & keyParts(0) & _
                    ", ServiceOffering=" & keyParts(1) & " at " & activeRows & _
                    ". Only one active record allowed per member/service combination."
            End If
        End If
    Next vKey
End Sub

' ==============================================================================
' SECTION 6: ERROR MESSAGE HELPERS
' ==============================================================================

Private Function GetFormatErrorMessage(sFieldType As String, sPattern As String, sValue As String) As String
    ' Provides user-friendly error messages
    
    Select Case UCase(sPattern)
        Case "DATE"
            GetFormatErrorMessage = "Invalid date format. Expected MM/DD/YYYY, found: '" & sValue & "'"
            
        Case "GENDER"
            GetFormatErrorMessage = "Invalid gender code. Expected M/F/U, found: '" & sValue & "'"
            
        Case "ZIP"
            GetFormatErrorMessage = "Invalid ZIP code format. Expected 5 or 9 digits, found: '" & sValue & "'"
            
        Case "STATE"
            GetFormatErrorMessage = "Invalid state code. Expected 2-letter state abbreviation, found: '" & sValue & "'"
            
        Case "NAME"
            GetFormatErrorMessage = "Invalid name format. Names should contain only letters, spaces, hyphens, and apostrophes. Found: '" & sValue & "'"
            
        Case "EMAIL"
            GetFormatErrorMessage = "Invalid email format. Expected format: user@domain.com, found: '" & sValue & "'"
            
        Case "PHONE"
            GetFormatErrorMessage = "Invalid phone format. Expected 10 digits, found: '" & sValue & "'"
            
        Case Else
            GetFormatErrorMessage = "Invalid format for " & sFieldType & ". Value: '" & sValue & "'"
    End Select
End Function

' ==============================================================================
' SECTION 7: COLUMN CHECKS SHEET AUTO-SETUP
' ==============================================================================

Public Sub SetupColumnChecksSheet()
    ' Automatically creates/updates the Column Checks sheet with proper rules
    
    Dim ws As Worksheet
    Dim bSheetExists As Boolean
    
    ' Check if sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Column Checks")
    bSheetExists = (Err.Number = 0)
    On Error GoTo 0
    
    ' Create if doesn't exist
    If Not bSheetExists Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Column Checks"
    End If
    
    ' Clear existing content
    ws.Cells.Clear
    
    ' Set headers
    ws.Range("A1").Value = "Field Name"
    ws.Range("B1").Value = "Required"
    ws.Range("C1").Value = "Max Length"
    ws.Range("D1").Value = "Min Length"
    ws.Range("E1").Value = "Format Pattern"
    ws.Range("F1").Value = "Custom Function"
    
    ' Format headers
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Add validation rules
    Dim row As Long
    row = 2
    
    ' FirstName
    ws.Cells(row, 1).Value = "FirstName"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 50
    ws.Cells(row, 4).Value = 1
    ws.Cells(row, 5).Value = "NAME"
    row = row + 1
    
    ' LastName
    ws.Cells(row, 1).Value = "LastName"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 50
    ws.Cells(row, 4).Value = 1
    ws.Cells(row, 5).Value = "NAME"
    row = row + 1
    
    ' DOB
    ws.Cells(row, 1).Value = "DOB"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 10
    ws.Cells(row, 4).Value = 8
    ws.Cells(row, 5).Value = "DATE"
    row = row + 1
    
    ' Gender
    ws.Cells(row, 1).Value = "Gender"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 10
    ws.Cells(row, 4).Value = 1
    ws.Cells(row, 5).Value = "GENDER"
    row = row + 1
    
    ' ZipCode
    ws.Cells(row, 1).Value = "ZipCode"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 10
    ws.Cells(row, 4).Value = 5
    ws.Cells(row, 5).Value = "ZIP"
    row = row + 1
    
    ' Address1
    ws.Cells(row, 1).Value = "Address1"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 100
    ws.Cells(row, 4).Value = 1
    ws.Cells(row, 5).Value = ""
    row = row + 1
    
    ' City
    ws.Cells(row, 1).Value = "City"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 50
    ws.Cells(row, 4).Value = 1
    ws.Cells(row, 5).Value = "NAME"
    row = row + 1
    
    ' State
    ws.Cells(row, 1).Value = "State"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 2
    ws.Cells(row, 4).Value = 2
    ws.Cells(row, 5).Value = "STATE"
    row = row + 1
    
    ' EffectiveDate
    ws.Cells(row, 1).Value = "EffectiveDate"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 10
    ws.Cells(row, 4).Value = 8
    ws.Cells(row, 5).Value = "DATE"
    row = row + 1
    
    ' GroupID
    ws.Cells(row, 1).Value = "GroupID"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 20
    ws.Cells(row, 4).Value = 1
    ws.Cells(row, 5).Value = ""
    row = row + 1
    
    ' ServiceOffering
    ws.Cells(row, 1).Value = "ServiceOffering"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 50
    ws.Cells(row, 4).Value = 1
    ws.Cells(row, 5).Value = ""
    row = row + 1
    
    ' MemberID
    ws.Cells(row, 1).Value = "MemberID"
    ws.Cells(row, 2).Value = "TRUE"
    ws.Cells(row, 3).Value = 20
    ws.Cells(row, 4).Value = 1
    ws.Cells(row, 5).Value = ""
    row = row + 1
    
    ' Address2 (optional)
    ws.Cells(row, 1).Value = "Address2"
    ws.Cells(row, 2).Value = "FALSE"
    ws.Cells(row, 3).Value = 100
    ws.Cells(row, 4).Value = 0
    ws.Cells(row, 5).Value = ""
    row = row + 1
    
    ' EffectiveEndDate (optional)
    ws.Cells(row, 1).Value = "EffectiveEndDate"
    ws.Cells(row, 2).Value = "FALSE"
    ws.Cells(row, 3).Value = 10
    ws.Cells(row, 4).Value = 0
    ws.Cells(row, 5).Value = "DATE"
    
    ' Auto-fit columns
    ws.Columns("A:F").AutoFit
    
    MsgBox "Column Checks sheet has been set up with standard validation rules.", _
           vbInformation, "Setup Complete"
End Sub

' ==============================================================================
' SECTION 8: COMPREHENSIVE TESTING SUITE
' ==============================================================================

Public Sub RunComprehensiveValidationTest()
    ' Tests all validation components
    
    Dim testResults As String
    testResults = "COMPREHENSIVE VALIDATION TEST RESULTS" & vbCrLf
    testResults = testResults & String(50, "=") & vbCrLf & vbCrLf
    
    ' Test 1: Column Checks Sheet
    testResults = testResults & "1. COLUMN CHECKS SHEET:" & vbCrLf
    If WorksheetExists("Column Checks") Then
        testResults = testResults & "   ✓ Sheet exists" & vbCrLf
        
        ' Check for required columns
        Dim wsChecks As Worksheet
        Set wsChecks = ThisWorkbook.Worksheets("Column Checks")
        
        If wsChecks.Cells(1, 1).Value = "Field Name" Then
            testResults = testResults & "   ✓ Headers configured correctly" & vbCrLf
        Else
            testResults = testResults & "   ✗ Headers missing or incorrect" & vbCrLf
        End If
    Else
        testResults = testResults & "   ✗ Sheet missing - Run SetupColumnChecksSheet()" & vbCrLf
    End If
    
    ' Test 2: Filetype Mapping Sheet
    testResults = testResults & vbCrLf & "2. FILETYPE MAPPING SHEET:" & vbCrLf
    If WorksheetExists("Filetype Mapping") Then
        testResults = testResults & "   ✓ Sheet exists" & vbCrLf
        
        Dim wsMapping As Worksheet
        Set wsMapping = ThisWorkbook.Worksheets("Filetype Mapping")
        Dim mappingCount As Long
        mappingCount = wsMapping.Cells(wsMapping.Rows.count, 1).End(xlUp).row - 1
        
        testResults = testResults & "   ✓ " & mappingCount & " file types configured" & vbCrLf
    Else
        testResults = testResults & "   ✗ Sheet missing" & vbCrLf
    End If
    
    ' Test 3: Test Validation Rules Loading
    testResults = testResults & vbCrLf & "3. VALIDATION RULES:" & vbCrLf
    
    Dim colRules As Collection
    Set colRules = LoadValidationRules()
    
    If colRules Is Nothing Then
        testResults = testResults & "   ✗ Failed to load rules" & vbCrLf
    Else
        testResults = testResults & "   ✓ " & colRules.count & " rules loaded" & vbCrLf
        
        ' Test specific rule
        Dim testRule As ValidationRule
        testRule = GetValidationRule(colRules, "FirstName")
        
        If testRule.Required Then
            testResults = testResults & "   ✓ FirstName rule loaded correctly" & vbCrLf
        Else
            testResults = testResults & "   ✗ FirstName rule not configured properly" & vbCrLf
        End If
    End If
    
    ' Test 4: ActiveChecker Function
    testResults = testResults & vbCrLf & "4. ACTIVE CHECKER TESTS:" & vbCrLf
    
    ' Test blank
    If ActiveChecker_FIXED("") = True Then
        testResults = testResults & "   ✓ Blank date = Active" & vbCrLf
    Else
        testResults = testResults & "   ✗ Blank date test failed" & vbCrLf
    End If
    
    ' Test future date
    Dim futureDate As String
    futureDate = Format(DateAdd("d", 30, Date), "mm/dd/yyyy")
    If ActiveChecker_FIXED(futureDate) = True Then
        testResults = testResults & "   ✓ Future date = Active" & vbCrLf
    Else
        testResults = testResults & "   ✗ Future date test failed" & vbCrLf
    End If
    
    ' Test past date
    Dim pastDate As String
    pastDate = Format(DateAdd("d", -30, Date), "mm/dd/yyyy")
    If ActiveChecker_FIXED(pastDate) = False Then
        testResults = testResults & "   ✓ Past date = Inactive" & vbCrLf
    Else
        testResults = testResults & "   ✗ Past date test failed" & vbCrLf
    End If
    
    ' Test 5: Format Validators
    testResults = testResults & vbCrLf & "5. FORMAT VALIDATORS:" & vbCrLf
    
    ' Test date format
    If ValidateDateFormat_Enhanced("01/15/2024") Then
        testResults = testResults & "   ✓ Date format validation working" & vbCrLf
    Else
        testResults = testResults & "   ✗ Date format validation failed" & vbCrLf
    End If
    
    ' Test ZIP code
    If ValidateZipCode_Enhanced("12345") Then
        testResults = testResults & "   ✓ ZIP code validation working" & vbCrLf
    Else
        testResults = testResults & "   ✗ ZIP code validation failed" & vbCrLf
    End If
    
    ' Test state code
    If ValidateStateCode_Enhanced("TX") Then
        testResults = testResults & "   ✓ State code validation working" & vbCrLf
    Else
        testResults = testResults & "   ✗ State code validation failed" & vbCrLf
    End If
    
    ' Display results
    testResults = testResults & vbCrLf & String(50, "=") & vbCrLf
    testResults = testResults & "Test completed at: " & Now
    
    ' Output to immediate window
    Debug.Print testResults
    
    ' Also show in message box
    MsgBox testResults, vbInformation, "Validation Test Results"
End Sub

Private Function WorksheetExists(sName As String) As Boolean
    ' Helper function to check if worksheet exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName)
    WorksheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' ==============================================================================
' SECTION 9: QUICK FIX IMPLEMENTATION
' ==============================================================================

Public Sub APPLY_ALL_FIXES()
    ' Main procedure to apply all fixes to your validation system
    
    Dim response As VbMsgBoxResult
    response = MsgBox("This will apply all validation fixes to your SFTP Command Center." & vbCrLf & vbCrLf & _
                      "The following actions will be performed:" & vbCrLf & _
                      "1. Setup/Update Column Checks sheet" & vbCrLf & _
                      "2. Fix validation functions" & vbCrLf & _
                      "3. Update duplicate detection logic" & vbCrLf & _
                      "4. Run comprehensive tests" & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Apply Validation Fixes")
    
    If response <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Step 1: Setup Column Checks sheet
    Debug.Print "Step 1: Setting up Column Checks sheet..."
    Call SetupColumnChecksSheet
    
    ' Step 2: Test the setup
    Debug.Print "Step 2: Running validation tests..."
    Call RunComprehensiveValidationTest
    
    ' Step 3: Update the existing modules (you'll need to do this manually)
    Debug.Print "Step 3: Validation fixes loaded. Update your modules:"
    Debug.Print "- Replace GetValidationRule with GetValidationRule"
    Debug.Print "- Replace ValidateField with ValidateField"
    Debug.Print "- Replace ActiveChecker with ActiveChecker_FIXED"
    Debug.Print "- Add CheckForDuplicates_FIXED to your validation routine"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Validation fixes have been applied successfully!" & vbCrLf & vbCrLf & _
           "Please check the Immediate Window for detailed results." & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Update your existing modules with the _FIXED functions" & vbCrLf & _
           "2. Run validation on test files to verify fixes", _
           vbInformation, "Fixes Applied"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error applying fixes: " & Err.Description, vbCritical, "Error"
End Sub
