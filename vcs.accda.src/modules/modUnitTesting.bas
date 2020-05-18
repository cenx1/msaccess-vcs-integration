Option Compare Database
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private Logger As New Logger
Private Sub log(sMessage As String, Optional logLevel As logLevel = info)
    If Logger.IsInitialized = False Then
        Logger.Initialize ("modUnitTesting")
    End If
    Logger.log sMessage, logLevel
End Sub
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


Sub t()
    log "T"
    Dim s As String * 1024
    log Len(s) & " length"
    Dim con As clsConcat
    Set con = New clsConcat
    Dim i As Integer, j As Integer
    
    For j = 0 To 1000
        log "-------------------------------"
        log "J IS " & j
        log "-------------------------------"
        For i = 65 To 90 ' 25 chars
            con.Add Chr(i)
            log con.GetStr
        Next i
    Next j
    
End Sub