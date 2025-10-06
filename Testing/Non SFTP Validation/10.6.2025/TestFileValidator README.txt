# CSV File Validation Tool

## Overview
This VBA-based tool provides automated validation for CSV files containing contact/address data. It performs comprehensive data quality checks on required and optional columns, generating detailed validation reports for data cleanup and quality assurance.

## Files Description

### FilePickerModule.bas
Contains the user interface for batch processing multiple CSV files.

**Main Function:** `RunValidationOnSelectedCSVs()`
- Opens a file picker dialog for selecting multiple CSV files
- Processes each file sequentially through the validation engine
- Handles file opening/closing operations safely
- Provides user feedback through message boxes

### ValidatorModule.bas
Contains the core validation logic and data quality checks.

**Main Function:** `Validate_WithRequiredColumnsSplitLog()`
- Validates data against predefined rules for each column
- Creates two separate validation logs (Required Columns and All Columns)
- Saves results as Excel files in a "Logs" subfolder

## How to Use

### Prerequisites
- Microsoft Excel with VBA enabled
- Write permissions to create a "Logs" folder in the workbook directory
- CSV files with headers in the first row

### Steps
1. Open Excel and enable macros
2. Import both VBA modules into your Excel workbook
3. Run the macro `RunValidationOnSelectedCSVs()`
4. Select one or more CSV files when prompted
5. Review the generated validation logs in the "Logs" folder

## Column Validation Rules

### Required Columns
The following columns **must** be present and cannot be blank:
- First Name
- Last Name
- Address Line 1
- City
- Zip Code
- E-mail Address

### Validation Checks by Column

#### First Name
- **Check Type:** Name Format
- **Max Length:** 50 characters
- **Valid Characters:** Letters (A-Z, a-z), Numbers (0-9), Spaces, Apostrophes ('), Hyphens (-)
- **Issues Flagged:**
  - Blank required field
  - Exceeds 50 characters
  - Invalid characters (anything not in allowed set)

#### Last Name
- **Check Type:** Name Format
- **Max Length:** 50 characters
- **Valid Characters:** Letters (A-Z, a-z), Numbers (0-9), Spaces, Apostrophes ('), Hyphens (-)
- **Issues Flagged:**
  - Blank required field
  - Exceeds 50 characters
  - Invalid characters (anything not in allowed set)

#### Address Line 1
- **Check Type:** Address Format
- **Max Length:** 150 characters
- **Valid Characters:** Letters (A-Z, a-z), Numbers (0-9), Spaces, Periods (.), Commas (,), Hyphens (-)
- **Issues Flagged:**
  - Blank required field
  - Exceeds 150 characters
  - Invalid characters (anything not in allowed set)

#### Address Line 2
- **Check Type:** Max Length Only
- **Max Length:** 150 characters
- **Valid Characters:** Any characters allowed
- **Issues Flagged:**
  - Exceeds 150 characters
- **Note:** This is an optional field (not required to have data)

#### City
- **Check Type:** Alphanumeric + Space
- **Max Length:** 150 characters
- **Valid Characters:** Letters (A-Z, a-z), Numbers (0-9), Spaces
- **Issues Flagged:**
  - Blank required field
  - Exceeds 150 characters
  - Invalid characters (anything not in allowed set)

#### Zip Code
- **Check Type:** Zip Format
- **Max Length:** 10 characters
- **Valid Format:** Must start with 5 digits, may have additional characters (supports ZIP+4 format)
- **Issues Flagged:**
  - Blank required field
  - Exceeds 10 characters
  - Invalid zip code format (doesn't start with 5 digits)

#### E-mail Address
- **Check Type:** Email Format
- **Max Length:** 150 characters
- **Valid Format:** Must contain both "@" and "." symbols
- **Issues Flagged:**
  - Blank required field
  - Exceeds 150 characters
  - Invalid email format (missing @ or . symbols)

## Validation Output

The tool generates two Excel log files for each validated CSV:

### Required Columns Log Sheet
- Contains validation issues only for the 6 required columns
- Focuses on critical data quality problems
- Use this for priority cleanup tasks

### All Columns Log Sheet
- Contains validation issues for all recognized columns (required and optional)
- Provides comprehensive data quality overview
- Use this for complete data cleanup

### Log File Format
Each log contains the following columns:
- **Row:** The row number in the original CSV where the issue was found
- **Column:** The column name where the issue occurred
- **Value:** The actual value that failed validation
- **Issue:** Description of the validation problem

### Log File Location
- Files are saved in a "Logs" subfolder within the workbook directory
- Naming convention: `[OriginalFileName]_ValidationLog.xlsx`

## Error Handling

- **Missing Required Columns:** Tool will only validate columns that exist in the CSV
- **File Access Issues:** Individual file errors won't stop batch processing
- **Invalid Data:** All validation issues are logged rather than stopping processing
- **Missing Logs Folder:** The tool will create the folder if it doesn't exist

## Common Validation Issues

### High Priority (Required Fields)
- Blank required fields (First Name, Last Name, Address Line 1, City, Zip Code, Email)
- Invalid email formats (missing @ or .)
- Invalid zip codes (not starting with 5 digits)

### Medium Priority (Format Issues)
- Names with special characters beyond allowed set
- Addresses with invalid punctuation
- City names with numbers or special characters

### Lower Priority (Length Issues)
- Fields exceeding maximum character limits
- Very long addresses or names that may cause system issues

## Best Practices

1. **Before Validation:** Ensure CSV files have proper headers in row 1
2. **After Validation:** Review Required Columns Log first for critical issues
3. **Data Cleanup:** Use the row numbers in logs to locate and fix issues in original data
4. **Regular Validation:** Run validation after any data imports or manual edits
5. **Archive Logs:** Keep validation logs for audit trails and data quality tracking