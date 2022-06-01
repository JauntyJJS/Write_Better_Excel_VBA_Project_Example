Attribute VB_Name = "Utilities"
Attribute VB_Description = "Functions used commonly in this project."
'@ModuleDescription("Functions used commonly in this project.")
Option Explicit
'@Folder("Utilities Functions")
'@Description("Get the file path of a folder selected by the user.")

'' Function: Get_Folder
'' --- Code
''  Public Function Get_Folder() As String
'' ---
''
'' Description:
''
'' Get the file path of a folder selected by the user.
''
'' The function will first open a pop up box to allow users
'' to select a folder.
''
'' (see Utilities_Get_Folder.png)
''
'' After the folder is selected, the file path of the folder is returned
''
'' Returns:
''    Returns the file path of a folder selected by the user
Public Function Get_Folder() As String
Attribute Get_Folder.VB_Description = "Get the file path of a folder selected by the user."
    'https://stackoverflow.com/questions/26392482/vba-excel-to-prompt-user-response-to-select-folder-and-return-the-path-as-string
    Dim Folder As FileDialog
    Dim Selected_Item As String
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    With Folder
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        Selected_Item = .SelectedItems.Item(1)
    End With
NextCode:
    Get_Folder = Selected_Item
    Set Folder = Nothing
End Function
