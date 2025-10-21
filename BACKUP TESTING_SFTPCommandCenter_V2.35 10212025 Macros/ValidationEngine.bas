Attribute VB_Name = "ValidationEngine"
Option Explicit

Type ColumnMapping
    fileType As String
    FirstName As Integer
    LastName As Integer
    DOB As Integer
    Gender As Integer
    ZipCode As Integer
    Address1 As Integer
    Address2 As Integer
    City As Integer
    State As Integer
    EffectiveDate As Integer
    serviceOffering As Integer
    memberID As Integer
    groupID As Integer
    effectiveEndDate As Integer
End Type

Public Type ValidationRule
    FieldType As String
    Required As Boolean
    MaxLength As Long
    MinLength As Long
    FormatPattern As String
    CustomFunction As String
End Type

Public Function GetColumnMapping(sFileType As String) As ColumnMapping
    Const sPROC_NAME As String = "GetColumnMapping"
    
    Dim oMapping As ColumnMapping
    Dim wsMapping As Worksheet
    Dim lLastRow As Long
    Dim lRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsMapping = ThisWorkbook.Worksheets("Filetype Mapping")
    lLastRow = wsMapping.Cells(wsMapping.Rows.count, "A").End(xlUp).row
    
    ' Initialize with empty values
    oMapping.fileType = ""
    oMapping.FirstName = 0
    oMapping.LastName = 0
    oMapping.DOB = 0
    oMapping.Gender = 0
    oMapping.ZipCode = 0
    oMapping.Address1 = 0
    oMapping.City = 0
    oMapping.State = 0
    oMapping.EffectiveDate = 0
    oMapping.groupID = 0
    oMapping.serviceOffering = 0
    oMapping.memberID = 0
    
    ' Find the FileType row
    For lRow = 2 To lLastRow
        If UCase(wsMapping.Cells(lRow, "A").Value) = UCase(sFileType) Then
            With oMapping
                .fileType = sFileType
                .FirstName = wsMapping.Cells(lRow, "B").Value
                .LastName = wsMapping.Cells(lRow, "C").Value
                .DOB = wsMapping.Cells(lRow, "D").Value
                .Gender = wsMapping.Cells(lRow, "E").Value
                .ZipCode = wsMapping.Cells(lRow, "F").Value
                .Address1 = wsMapping.Cells(lRow, "G").Value
                .City = wsMapping.Cells(lRow, "H").Value
                .State = wsMapping.Cells(lRow, "I").Value
                .EffectiveDate = wsMapping.Cells(lRow, "J").Value
                .groupID = wsMapping.Cells(lRow, "K").Value
                .serviceOffering = wsMapping.Cells(lRow, "L").Value
                .memberID = wsMapping.Cells(lRow, "M").Value
                .effectiveEndDate = wsMapping.Cells(lRow, "N").Value
            End With
            
            ' Return the Type directly (no Set keyword)
            GetColumnMapping = oMapping
            Exit Function
        End If
    Next lRow
    
    ' FileType not found - return empty mapping
    GetColumnMapping = oMapping
    Exit Function
    
ErrorHandler:
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description, sFileType)
    GetColumnMapping = oMapping
End Function

Public Function LoadValidationRules() As Collection
    Const sPROC_NAME As String = "LoadValidationRules"
    
    Dim colRules As New Collection
    Dim wsRules As Worksheet
    Dim lRow As Long
    Dim lLastRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Reference the Column Checks sheet
    Set wsRules = ThisWorkbook.Worksheets("Column Checks")
    lLastRow = wsRules.Cells(wsRules.Rows.count, "A").End(xlUp).row
    
    ' Expected Column Layout in "Column Checks" sheet:
    ' Column A: Field Name (e.g., FirstName, LastName, DOB, Gender, ZipCode, Address1, City, State, EffectiveDate, ServiceOffering, etc.)
    ' Column B: Required (TRUE/FALSE)
    ' Column C: Max Length
    ' Column D: Min Length
    ' Column E: Format Pattern (regex or format type)
    ' Column F: Custom Function (optional)
    
    ' Process each row (assuming row 1 has headers)
    For lRow = 2 To lLastRow
        Dim sFieldType As String
        Dim sRuleData As String
        
        sFieldType = Trim(wsRules.Cells(lRow, "A").Value)
        
        If sFieldType <> "" Then
            ' Create a delimited string with rule data
            sRuleData = sFieldType & "|" & _
                       wsRules.Cells(lRow, "B").Value & "|" & _
                       wsRules.Cells(lRow, "C").Value & "|" & _
                       wsRules.Cells(lRow, "D").Value & "|" & _
                       wsRules.Cells(lRow, "E").Value & "|" & _
                       wsRules.Cells(lRow, "F").Value
            
            ' Add the string data to collection using FieldType as key
            On Error Resume Next
            colRules.Add sRuleData, sFieldType
            If Err.Number <> 0 Then
                Debug.Print "Duplicate field type found: " & sFieldType
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            ' Log the loaded rule
            Debug.Print "Loaded rule for: " & sFieldType & " - Required: " & wsRules.Cells(lRow, "B").Value & _
                       ", MaxLen: " & wsRules.Cells(lRow, "C").Value & ", MinLen: " & wsRules.Cells(lRow, "D").Value
        End If
    Next lRow
    
    Debug.Print "Total validation rules loaded: " & colRules.count
    
    Set LoadValidationRules = colRules
    Exit Function
    
ErrorHandler:
    MsgBox "Error loading validation rules from 'Column Checks' sheet: " & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "LoadValidationRules Error"
    
    If colRules Is Nothing Then Set colRules = New Collection
    Set LoadValidationRules = colRules
End Function

Private Function FindValidationRule(colRules As Collection, sFieldType As String) As ValidationRule
    Dim oRule As ValidationRule
    Dim emptyRule As ValidationRule
    Dim sRuleData As String
    Dim vParts As Variant
    
    ' Initialize empty rule
    emptyRule.FieldType = ""
    emptyRule.Required = False
    emptyRule.MaxLength = 0
    emptyRule.MinLength = 0
    emptyRule.FormatPattern = ""
    emptyRule.CustomFunction = ""
    
    On Error GoTo NotFound
    
    ' Get the rule data string from collection
    sRuleData = colRules.Item(sFieldType)
    
    ' Parse the delimited string back into ValidationRule Type
    vParts = Split(sRuleData, "|")
    
    If UBound(vParts) >= 5 Then
        oRule.FieldType = vParts(0)
        oRule.Required = (UCase(vParts(1)) = "Y" Or UCase(vParts(1)) = "TRUE" Or UCase(vParts(1)) = "y")
        oRule.MaxLength = Val(vParts(2))
        oRule.MinLength = Val(vParts(3))
        oRule.FormatPattern = vParts(4)
        oRule.CustomFunction = vParts(5)
        
        FindValidationRule = oRule
        Exit Function
    End If
    
NotFound:
    ' Return empty rule if not found or error
    FindValidationRule = emptyRule
End Function

Private Function ValidateFieldFormat(sValue As String, sFieldType As String, sPattern As String) As Boolean
    Select Case UCase(sFieldType)
        Case "DOB", "EFFECTIVEDATE"
            ValidateFieldFormat = ValidateDateFormat(sValue)
        Case "GENDER"
            ValidateFieldFormat = ValidateGenderCode(sValue)
        Case "ZIPCODE"
            ValidateFieldFormat = ValidateZipCode(sValue)
        Case "FIRSTNAME", "LASTNAME", "CITY"
            ValidateFieldFormat = ValidateNameFormat(sValue)
        Case "STATE"
            ValidateFieldFormat = ValidateStateCode(sValue)
        Case Else
            ' Use regex pattern if provided
            If sPattern <> "" Then
                ValidateFieldFormat = ValidateWithRegex(sValue, sPattern)
            Else
                ValidateFieldFormat = True ' No specific validation
            End If
        
    End Select
End Function

Private Function ValidateDateFormat(sValue As String) As Boolean
    On Error Resume Next
    Dim dtTest As Date
    dtTest = CDate(sValue)
    ValidateDateFormat = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function ValidateGenderCode(sValue As String) As Boolean
    Dim vValidCodes As Variant
    vValidCodes = Array("M", "F", "MALE", "FEMALE", "1", "2", "0", "U", "UNKNOWN")
    
    Dim i As Long
    For i = 0 To UBound(vValidCodes)
        If UCase(Trim(sValue)) = UCase(CStr(vValidCodes(i))) Then
            ValidateGenderCode = True
            Exit Function
        End If
    Next i
    
    ValidateGenderCode = False
End Function

Private Function ValidateZipCode(sValue As String) As Boolean
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    ' US: 12345 or 12345-6789
    oRegex.pattern = "^\d{5}(-\d{4})?$"
    ValidateZipCode = oRegex.Test(sValue)
End Function

Private Function ValidateNameFormat(sValue As String) As Boolean
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    ' Allow letters, spaces, hyphens, apostrophes, periods
    oRegex.pattern = "^[a-zA-Z][a-zA-Z\s\-'\.]{1,49}$"
    oRegex.IgnoreCase = True
    
    ValidateNameFormat = oRegex.Test(Trim(sValue)) And Len(Trim(sValue)) >= 2
End Function

Private Function ValidateStateCode(sValue As String) As Boolean
    ' This could be expanded with a full list of state codes
    ValidateStateCode = (Len(Trim(sValue)) = 2)
End Function

Private Function ValidateWithRegex(sValue As String, sPattern As String) As Boolean
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    oRegex.pattern = sPattern
    ValidateWithRegex = oRegex.Test(sValue)
End Function
