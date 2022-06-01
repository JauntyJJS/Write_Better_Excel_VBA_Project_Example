Attribute VB_Name = "Sample_Type_Identifier"
Attribute VB_Description = "Identify the sample type for a given input string."
'@ModuleDescription("Identify the sample type for a given input string.")
Option Explicit

'@Folder("QC_Sample_Type_Identification")
'@Description("Check is the input string (sample name) is an EQC.")

'' Function: Is_EQC
'' --- Code
''  Public Function Is_EQC(ByVal FileName As String) As Boolean
'' ---
''
'' Description:
''
'' Check is the input string (sample name) is an EQC.
''
'' Parameters:
''
''    FileName - Input string to check if it is an EQC
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "EQC", "Eqc", "eqc"
''
'' Examples:
''
'' --- Code
''    Dim EQCTestArray As Variant
''    Dim arrayIndex As Integer
''
''    EQCTestArray = Array("EQC", "001_EQC_TQC prerun 01")
''
''    For arrayIndex = 0 To UBound(EQCTestArray) - LBound(EQCTestArray)
''        Debug.Print Sample_Type_Identifier.Is_EQC(CStr(EQCTestArray(arrayIndex))) & ": " & _
''                    EQCTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_EQC(ByVal FileName As String) As Boolean
Attribute Is_EQC.VB_Description = "Check is the input string (sample name) is an EQC."
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim EQCRegEx As RegExp
    Set EQCRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    EQCRegEx.Pattern = "(EQC|[Ee]qc)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_EQC = EQCRegEx.Test(OnlyLettersText)
    
End Function
