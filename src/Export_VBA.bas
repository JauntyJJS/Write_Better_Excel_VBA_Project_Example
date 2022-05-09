Attribute VB_Name = "Export_VBA"
Option Explicit
'@Folder("Export VBA Source Code Functions")
'@IgnoreModule IntegerDataType

'' Function: Export_Visual_Basic_Code
''
'' Description:
''
'' Export all VBA source codes to a user selected folder.
''
'' The function will first open a pop up box to allow users
'' to select a folder.
''
'' (see Utilities_Get_Folder.png)
''
'' After the folder is selected, the VBA source codes
'' will be exported to that folder.
''
'' (see Export_VBA_Output.png)
''
'' Do note that files will be overwritten without warning.
''
Public Sub Export_Visual_Basic_Code()
    'https://gist.github.com/steve-jansen/7589478
    Const Module As Integer = 1
    Const ClassModule As Integer = 2
    Const Form As Integer = 3
    Const Document As Integer = 100
    'Const Padding As Integer = 24
    
    Dim VBComponent As Object
    Dim Count As Integer
    Dim Path As String
    Dim Directory As String
    Dim Extension As String
    
    ' Folder must exist prior to outputting the source file
    Directory = Utilities.Get_Folder()
    If Directory = vbNullString Then
        Exit Sub
    End If
    
    Count = 0
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                Extension = ".cls"
            Case Form
                Extension = ".frm"
            Case Module
                Extension = ".bas"
            Case Else
                Extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        Path = Directory & "\" & VBComponent.Name & Extension
        VBComponent.Export Path
        
        If Err.Number <> 0 Then
            MsgBox "Failed to export " & VBComponent.Name & " to " & Path, vbCritical
        Else
            Count = Count + 1
            'Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    MsgBox "Successfully exported " & CStr(Count) & " VBA files to " & Directory

End Sub

