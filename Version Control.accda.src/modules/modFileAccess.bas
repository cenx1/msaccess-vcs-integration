Option Explicit
Option Compare Database
Option Private Module


Private Declare PtrSafe Function getTempPath Lib "kernel32" Alias "GetTempPathA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
    
Private Declare PtrSafe Function getTempFileName Lib "kernel32" Alias "GetTempFileNameA" ( _
    ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long

''' Maps a character string to a UTF-16 (wide character) string
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
    ) As Long

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long


' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

' Cache the Ucs2 requirement for this database
Private m_blnUcs2 As Boolean
Private m_strDbPath As String


'---------------------------------------------------------------------------------------
' Procedure : RequiresUcs2
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Returns true if the current database requires objects to be converted
'           : to Ucs2 format before importing. (Caching value for subsequent calls.)
'           : While this involves creating a new querydef object each time, the idea
'           : is that this would be faster than exporting a form if no queries exist
'           : in the current database.
'---------------------------------------------------------------------------------------
'
Public Function RequiresUcs2(Optional blnUseCache As Boolean = True) As Boolean

    Dim strTempFile As String
    Dim frm As Access.Form
    Dim strName As String
    Dim dbs As DAO.Database
    
    ' See if we already have a cached value
    If (m_strDbPath <> CurrentProject.FullName) Or Not blnUseCache Then
    
        ' Get temp file name
        strTempFile = GetTempFile
        
        ' Can't create querydef objects in ADP databases, so we have to use something else.
        If CurrentProject.ProjectType = acADP Then
            ' Create and export a blank form object.
            ' Turn of screen updates to improve performance and avoid flash.
            DoCmd.Echo False
            'strName = "frmTEMP_UCS2_" & Round(Timer)
            Set frm = Application.CreateForm
            strName = frm.Name
            DoCmd.Close acForm, strName, acSaveYes
            Application.SaveAsText acForm, strName, strTempFile
            DoCmd.DeleteObject acForm, strName
            DoCmd.Echo True
        Else
            ' Standard MDB database.
            ' Create and export a querydef object. Fast and light.
            strName = "qryTEMP_UCS2_" & Round(Timer)
            Set dbs = CurrentDb
            dbs.CreateQueryDef strName, "SELECT 1"
            Application.SaveAsText acQuery, strName, strTempFile
            dbs.QueryDefs.Delete strName
        End If
        
        ' Test and delete temp file
        m_strDbPath = CurrentProject.FullName
        m_blnUcs2 = HasUcs2Bom(strTempFile)
        FSO.DeleteFile strTempFile, True

    End If

    ' Return cached value
    RequiresUcs2 = m_blnUcs2
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ConvertUcs2Utf8
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Convert a UCS2-little-endian encoded file to UTF-8.
'           : Typically the source file will be a temp file.
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUcs2Utf8(strSourceFile As String, strDestinationFile As String, _
    Optional blnDeleteSourceFileAfterConversion As Boolean = True)

    Dim strText As String
    Dim utf8Bytes() As Byte
    Dim fnum As Integer
    Dim blnIsAdp As Boolean
    Dim intTristate As Tristate
    
    ' Make sure the path exists and remove any existing file.
    VerifyPath FSO.GetParentFolderName(strDestinationFile)
    If FSO.FileExists(strDestinationFile) Then FSO.DeleteFile strDestinationFile, True
    
    ' ADP Projects do not use the UCS BOM, but may contain mixed UTF-16 content
    ' representing unicode characters.
    blnIsAdp = (CurrentProject.ProjectType = acADP)
    
    ' Check the first couple characters in the file for a UCS BOM.
    If HasUcs2Bom(strSourceFile) Or blnIsAdp Then
    
        ' Determine format
        If blnIsAdp Then
            ' Possible mixed UTF-16 content
            intTristate = TristateMixed
        Else
            ' Fully encoded as UTF-16
            intTristate = TristateTrue
        End If
            
        ' Read file contents and delete (temp) source file
        With FSO.OpenTextFile(strSourceFile, ForReading, False, intTristate)
            strText = .ReadAll
            .Close
        End With
        
        ' Write as UTF-8 in the destination file.
        WriteFile strText, strDestinationFile
        
        ' Remove the source (temp) file if specified
        If blnDeleteSourceFileAfterConversion Then FSO.DeleteFile strSourceFile, True
    Else
        ' No conversion needed, move/copy to destination.
        If blnDeleteSourceFileAfterConversion Then
            FSO.MoveFile strSourceFile, strDestinationFile
        Else
            FSO.CopyFile strSourceFile, strDestinationFile
        End If
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertUtf8Ucs2
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Convert the file to old UCS-2 unicode format.
'           : Typically the destination file will be a temp file.
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUtf8Ucs2(strSourceFile As String, strDestinationFile As String, _
    Optional blnDeleteSourceFileAfterConversion As Boolean = True)

    Dim strText As String
    Dim utf8Bytes() As Byte
    Dim fnum As Integer

    ' Make sure the path exists before we write a file.
    VerifyPath FSO.GetParentFolderName(strDestinationFile)
    
    If FSO.FileExists(strDestinationFile) Then FSO.DeleteFile strDestinationFile, True
    
    If HasUcs2Bom(strSourceFile) Then
        ' No conversion needed, move/copy to destination.
        If blnDeleteSourceFileAfterConversion Then
            FSO.MoveFile strSourceFile, strDestinationFile
        Else
            FSO.CopyFile strSourceFile, strDestinationFile
        End If
    Else
        ' Read file contents and convert byte array to string
        utf8Bytes = GetFileBytes(strSourceFile)
        strText = Utf8BytesToString(utf8Bytes)
        
        ' Write as UCS-2 LE (BOM)
        With FSO.CreateTextFile(strDestinationFile, True, TristateTrue)
            .Write strText
            .Close
        End With
        
        ' Remove original file if specified.
        If blnDeleteSourceFileAfterConversion Then FSO.DeleteFile strSourceFile, True
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasUtf8Bom
' Author    : Adam Waller
' Date      : 7/30/2020
' Purpose   : Returns true if the file begins with a UTF-8 BOM
'---------------------------------------------------------------------------------------
'
Public Function HasUtf8Bom(strFilePath As String) As Boolean
    HasUtf8Bom = FileHasBom(strFilePath, UTF8_BOM)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasUcs2Bom
' Author    : Adam Waller
' Date      : 8/1/2020
' Purpose   : Returns true if the file begins with
'---------------------------------------------------------------------------------------
'
Public Function HasUcs2Bom(strFilePath As String) As Boolean
    HasUcs2Bom = FileHasBom(strFilePath, UCS2_BOM)
End Function


'---------------------------------------------------------------------------------------
' Procedure : FileHasBom
' Author    : Adam Waller
' Date      : 8/1/2020
' Purpose   : Check for the specified BOM
'---------------------------------------------------------------------------------------
'
Private Function FileHasBom(strFilePath As String, strBom As String) As Boolean
    Dim strFound As String
    strFound = StrConv((GetFileBytes(strFilePath, Len(strBom))), vbUnicode)
    FileHasBom = (strFound = strBom)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveUTF8BOM
' Author    : Adam Kauffman
' Date      : 1/24/2019
' Purpose   : Will remove a UTF8 BOM from the start of the string passed in.
'---------------------------------------------------------------------------------------
'
Public Function RemoveUTF8BOM(ByVal fileContents As String) As String
    Dim UTF8BOM As String
    UTF8BOM = Chr$(239) & Chr$(187) & Chr$(191) ' == &HEFBBBF
    Dim fileBOM As String
    fileBOM = Left$(fileContents, 3)
    
    If fileBOM = UTF8BOM Then
        RemoveUTF8BOM = Mid$(fileContents, 4)
    Else ' No BOM detected
        RemoveUTF8BOM = fileContents
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTempFile
' Author    : Adapted by Adam Waller
' Date      : 1/23/2019
' Purpose   : Generate Random / Unique temporary file name.
'---------------------------------------------------------------------------------------
'
Public Function GetTempFile(Optional strPrefix As String = "VBA") As String

    Dim strPath As String * 512
    Dim strName As String * 576
    Dim lngReturn As Long
    
    lngReturn = getTempPath(512, strPath)
    lngReturn = getTempFileName(strPath, strPrefix, 0, strName)
    If lngReturn <> 0 Then GetTempFile = Left$(strName, InStr(strName, vbNullChar) - 1)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : BytesLength
' Author    : Casper Englund
' Date      : 2020/05/01
' Purpose   : Return length of byte array
'---------------------------------------------------------------------------------------
Private Function BytesLength(abBytes() As Byte) As Long
    
    ' Ignore error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Utf8BytesToString
' Author    : Adapted by Casper Englund
' Date      : 2020/05/01
' Purpose   : Return VBA "Unicode" string from byte array encoded in UTF-8
'---------------------------------------------------------------------------------------
Public Function Utf8BytesToString(abUtf8Array() As Byte) As String
    
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Dim bUtf8Bom As Boolean
    
    Utf8BytesToString = vbNullString
    
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    bUtf8Bom = abUtf8Array(0) = 239 _
      And abUtf8Array(1) = 187 _
      And abUtf8Array(2) = 191
    
    If bUtf8Bom Then
        Dim i As Long
        Dim abTempArr() As Byte
        ReDim abTempArr(BytesLength(abUtf8Array) - 3)
        For i = 3 To UBound(abUtf8Array)
            abTempArr(i - 3) = abUtf8Array(i)
        Next i
        
        abUtf8Array = abTempArr
    End If
    
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)

End Function


'---------------------------------------------------------------------------------------
' Procedure : Utf8BytesFromString
' Author    : Adapted by Casper Englund
' Date      : 2020/05/01
' Purpose   : Return byte array with VBA "Unicode" string encoded in UTF-8
'---------------------------------------------------------------------------------------
Public Function Utf8BytesFromString(strInput As String) As Byte()

    Dim nBytes As Long
    Dim abBuffer() As Byte
    
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
    
End Function