Attribute VB_Name = "Export_VBA"

' Requires the following to be enabled in the workbook to run
' Microsoft Scripting Runtime
' Microsoft Visual Basic for Applications Extensibility 5.3
Sub AddComponent()
    On Error Resume Next
    Dim MSR, VBAE As String
    MSR = "C:\Windows\System32\scrrun.dll" ' Path to Microsoft Scripting Runtime
    VBAE = "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB" ' Path to Microsoft Visual Basic for Applications Extensibility 5.3

    With ThisWorkbook.VBProject.References
        .AddFromFile (MSR)
        .AddFromFile (VBAE)
    End With
    
End Sub
Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.bas"
        Kill FolderWithVBAProjectFiles & "\*.frm"
        Kill FolderWithVBAProjectFiles & "\*.frx"
        Kill FolderWithVBAProjectFiles & "\*.cls"
        Kill FolderWithVBAProjectFiles & "\*.txt"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                'bExport = False
                
                szFileName = szFileName & ".cls"

        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready - " & szExportPath
End Sub


Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = ThisWorkbook.Path

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If

    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    

    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    ElseIf FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        SpecialPath = Environ("Userprofile") & "\Downloads\"
        If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
            On Error Resume Next
            MkDir SpecialPath & "VBAProjectFiles"
            On Error GoTo 0
        End If
        If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
            FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
        End If
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
        Dim vbProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set vbProj = ActiveWorkbook.VBProject
        
        For Each VBComp In vbProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                vbProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function
Sub ActivateReferenceLibrary()

'PURPOSE: Show How To Activate Specific Object Libraries
'SOURCE: www.TheSpreadsheetGuru.com

'Error Handler in Case Reference is Already Activated
  On Error Resume Next
    
    'Activate Microsoft Scripting Runtime
        ThisWorkbook.VBProject.References.AddFromGuid _
          GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", _
          Major:=0, Minor:=0
    
    'Activate Visual Basic for Applications Extensibility Library (version 5.3)
      ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{0002E157-0000-0000-C000-000000000046}", _
        Major:=0, Minor:=0 'Use zeroes to default to latest version

'Reset Error Handler
  On Error GoTo 0
  
  ExportModules
  
End Sub
Sub Display_GUID_Info()

'PURPOSE: Displays GUID information for each active _
Object Library reference in the VBA project
'SOURCE: www.TheSpreadsheetGuru.com

'Dim ref As Reference
    
'Loop Through Each Active Reference (Displays in Immediate Window [ctrl + g])
  For Each ref In ThisWorkbook.VBProject.References
    Debug.Print "Reference Name: ", ref.Name
    Debug.Print "Path: ", ref.FullPath
    Debug.Print "GUID: " & ref.GUID
    Debug.Print "Version: " & ref.Major & "." & ref.Minor
    Debug.Print " "
  Next ref
  
End Sub


