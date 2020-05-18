Option Explicit
Option Compare Database
Option Private Module


'---------------------------------------------------------------------------------------
' Procedure : ExportReferences
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Export references to a csv text file
'---------------------------------------------------------------------------------------
'
Public Sub ExportReferences(strFolder As String, cOptions As clsOptions)
    
    Dim cData As New clsConcat
    Dim ref As Reference
    Dim intCnt As Integer
    
    For Each ref In Application.References
        If ref.GUID = "" Then ' references of types mdb,accdb,mde etc don't have a GUID
            cData.Add ref.FullPath
            cData.Add vbCrLf
            log "  [" & ref.Name & "]", cOptions.ShowDebug
       Else
            With cData
                .Add ref.GUID
                .Add ","
                .Add ref.Name
                .Add ","
                .Add CStr(ref.Major)
                .Add ","
                .Add CStr(ref.Minor)
                .Add vbCrLf
            End With
            log "  " & ref.Name & " " & ref.Major & "." & ref.Minor, cOptions.ShowDebug
        End If
        intCnt = intCnt + 1
    Next ref
    
    ' Write ouput to file
    WriteFile cData.GetStr, strFolder & "references.csv"

    ' Show summary
    If cOptions.ShowDebug Then
        log "[" & intCnt & "] references exported."
    Else
        log "[" & intCnt & "]"
    End If
    
End Sub


' Import References from a CSV, true=SUCCESS
Public Function ImportReferences(obj_path As String) As Boolean
    
    Dim InFile As Scripting.TextStream
    Dim line As String
    Dim Item() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim fileName As String
    Dim refName As String
    fileName = Dir(obj_path & "references.csv")
    If Len(fileName) = 0 Then
        ImportReferences = False
        Exit Function
    End If
    Set InFile = FSO.OpenTextFile(obj_path & fileName, ForReading)
On Error GoTo failed_guid
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        Item = Split(line, ",")
        If UBound(Item) = 2 Then 'a ref with a guid
          GUID = Trim(Item(0))
          Major = CLng(Item(1))
          Minor = CLng(Item(2))
          Application.References.AddFromGuid GUID, Major, Minor
        Else
          refName = Trim(Item(0))
          Application.References.AddFromFile refName
        End If
go_on:
    Loop
On Error GoTo 0
    InFile.Close
    Set InFile = Nothing
    ImportReferences = True
    Exit Function
failed_guid:
    If Err.Number = 32813 Then
        'The reference is already present in the access project - so we can ignore the error
        Resume Next
    Else
        MsgBox "Failed to register " & GUID, , "Error: " & Err.Number
        'Do we really want to carry on the import with missing references??? - Surely this is fatal
        Resume go_on
    End If
    
End Function