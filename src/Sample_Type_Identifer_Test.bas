Attribute VB_Name = "Sample_Type_Identifer_Test"
Attribute VB_Description = "Test units for Sample_Type_Identifier."
'@ModuleDescription("Test units for Sample_Type_Identifier.")
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
'Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Test if Sample_Type_Identifier.Is_EQC is working")

'' Function: Is_EQC_Test
'' --- Code
''  Public Sub Is_EQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_EQC is working
''
'' Test data are
''
''  - A string array EQCTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_EQC gives
'' True to all entries in EQCTestArray
''
Public Sub Is_EQC_Test()
Attribute Is_EQC_Test.VB_Description = "Test if Sample_Type_Identifier.Is_EQC is working"
    On Error GoTo TestFail
    
    Dim EQCTestArray As Variant
    Dim arrayIndex As Integer
    
    EQCTestArray = Array("EQC", "001_EQC_TQC prerun 01")
           
    For arrayIndex = 0 To UBound(EQCTestArray) - LBound(EQCTestArray)
        'Debug.Print Sample_Type_Identifier.Is_EQC(CStr(EQCTestArray(arrayIndex))) & ": " & _
                     EQCTestArray(arrayIndex)
        MsgBox Sample_Type_Identifier.Is_EQC(CStr(EQCTestArray(arrayIndex))) & ": " & _
                     EQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_EQC(CStr(EQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
