Option Compare Database
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

' Test shows that UCS-2 files exported by Access make round trip through our conversions.
'@TestMethod("TextConversions")
Public Sub TestUCS2toUTF8RoundTrip()
    On Error GoTo TestFail
    
    'Arrange:
    Dim queryName As String
    queryName = "Temp_Test_Query_Delete_Me"
    Dim tempFileName As String
    tempFileName = GetTempFile()
    
    Dim UCStoUCS As String
    Dim UCStoUTF As String
    Dim UTFtoUTF As String
    Dim UTFtoUCS As String
    UCStoUCS = tempFileName & "UCS-2toUCS-2"
    UCStoUTF = tempFileName & "UCS-2toUTF-8"
    UTFtoUTF = tempFileName & "UTF-8toUTF-8"
    UTFtoUCS = tempFileName & "UTF-8toUCS-2"
    
    ' Use temporary query to export example file
    CurrentDb.CreateQueryDef queryName, "SELECT * FROM TEST WHERE TESTING=TRUE"
    Application.SaveAsText acQuery, queryName, tempFileName
    CurrentDb.QueryDefs.Delete queryName
        
    'Act:
    ConvertUtf8Ucs2 tempFileName, UCStoUCS
    ConvertUcs2Utf8 UCStoUCS, UCStoUTF
    ConvertUcs2Utf8 UCStoUTF, UTFtoUTF
    ConvertUtf8Ucs2 UTFtoUTF, UTFtoUCS
    
    ' Read original export
    Dim originalExport As String
    With FSO.OpenTextFile(tempFileName, , , TristateTrue)
        originalExport = .ReadAll
        .Close
    End With
    
    ' Read final file that went through all permutations of conversion
    Dim finalFile As String
    With FSO.OpenTextFile(UTFtoUCS, , , TristateTrue)
        finalFile = .ReadAll
        .Close
    End With
    
    ' Cleanup temp files
    FSO.DeleteFile tempFileName
    FSO.DeleteFile UCStoUCS
    FSO.DeleteFile UCStoUTF
    FSO.DeleteFile UTFtoUTF
    FSO.DeleteFile UTFtoUCS
    
    'Assert:
    Assert.AreEqual originalExport, finalFile
    
    GoTo TestExit
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

TestExit:
    
End Sub

'@TestMethod("TextConversions")
Public Sub TestUCS2toUTF8RoundTripNewVersion()
    On Error GoTo TestFail
    
    'Arrange:
    Dim queryName As String
    queryName = "Temp_Test_Query_Delete_Me"
    Dim tempFileName As String
    tempFileName = GetTempFile()
    Debug.Print tempFileName
    
    Dim UCStoUCS As String
    Dim UCStoUTF As String
    Dim UTFtoUTF As String
    Dim UTFtoUCS As String
    UCStoUCS = tempFileName & "UCS-2toUCS-2"
    UCStoUTF = tempFileName & "UCS-2toUTF-8"
    UTFtoUTF = tempFileName & "UTF-8toUTF-8"
    UTFtoUCS = tempFileName & "UTF-8toUCS-2"
    
    ' Use temporary query to export example file
    Debug.Print "Query name: " & queryName
    CurrentDb.CreateQueryDef queryName, "SELECT * FROM TEST WHERE TESTING=TRUE AND SUPPORTED='∆ÿ≈'"
    Debug.Print "Save to temp file: " & tempFileName
    Application.SaveAsText acQuery, queryName, tempFileName
    Debug.Print "Delete query name: " & queryName
    CurrentDb.QueryDefs.Delete queryName
        
    'Act:
    Debug.Print "Act"
    Debug.Print "Filename: " & tempFileName
    ConvertUtf8Ucs2 tempFileName, UCStoUCS
    Debug.Print "Filename: " & UCStoUCS
    ConvertUcs2Utf8 UCStoUCS, UCStoUTF
    Debug.Print "Filename: " & UCStoUTF
    ConvertUcs2Utf8 UCStoUTF, UTFtoUTF
    Debug.Print "Filename: " & UTFtoUTF
    ConvertUtf8Ucs2 UTFtoUTF, UTFtoUCS
    Debug.Print "Filename: " & UTFtoUCS
    
    ' Read original export
    Debug.Print "Read original export"
    Dim originalExport As String
    With FSO.OpenTextFile(tempFileName, , , TristateTrue)
        originalExport = .ReadAll
        .Close
    End With
    
    ' Read final file that went through all permutations of conversion
    Debug.Print "Read final file"
    Dim finalFile As String
    With FSO.OpenTextFile(UTFtoUCS, , , TristateTrue)
        finalFile = .ReadAll
        .Close
    End With
    
    ' Cleanup temp files
    Debug.Print "Clean up"
    FSO.DeleteFile tempFileName
    FSO.DeleteFile UTFtoUTF
    FSO.DeleteFile UTFtoUCS
    
    'Assert:
    Debug.Print "Original Export: " & vbCrLf & originalExport
    Debug.Print "Final File: " & vbCrLf & finalFile
    Assert.AreEqual originalExport, finalFile
    
    GoTo TestExit
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

TestExit:
    
End Sub