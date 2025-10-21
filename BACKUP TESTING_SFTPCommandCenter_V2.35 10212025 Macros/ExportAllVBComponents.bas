Attribute VB_Name = "ExportAllVBComponents"
Option Explicit

Sub ExportAllVBComponents()
    Dim VBComp As VBIDE.VBComponent
    Dim FolderPath As String
    Dim SubFolder As String
    Dim ProjectName As String
    Dim ExportPath As String
    Dim Ext As String
    Dim fDialog As fileDialog
    Dim todaysDate As String

    ' Ensure the VBA Extensibility library is enabled:
    ' Go to Tools > References > Check "Microsoft Visual Basic for Applications Extensibility 5.3"
    
    ' Get current project name
    ProjectName = ThisWorkbook.Name
    ProjectName = Left(ProjectName, InStrRev(ProjectName, ".") - 1)
    
    ' Format current date as mmddyyyy
    todaysDate = Format(Date, "mmddyyyy")
    
    ' Prompt user to select destination folder
    Set fDialog = Application.fileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select Folder to Export Project Modules"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        FolderPath = .SelectedItems(1)
    End With
    
    ' Create subfolder with project name and date
    SubFolder = FolderPath & "\" & ProjectName & " " & todaysDate & " Macros"
    If Dir(SubFolder, vbDirectory) = "" Then MkDir SubFolder
    
    ' Export all VB components
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                Ext = ".cls"
            Case vbext_ct_StdModule
                Ext = ".bas"
            Case vbext_ct_MSForm
                Ext = ".frm"
            Case Else
                Ext = ""
        End Select
        
        If Ext <> "" Then
            VBComp.Export SubFolder & "\" & VBComp.Name & Ext
        End If
    Next VBComp
    
    ' Confirm completion
    MsgBox "All modules, class modules, and forms have been exported to:" & vbCrLf & SubFolder, vbInformation, "Export Complete"

End Sub

