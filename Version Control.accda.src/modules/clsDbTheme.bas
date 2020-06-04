Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


Private m_AllItems As Collection
Private m_Dbs As DAO.Database

' This is used to pass a reference to the record back
' into the class for loading the private variables
' with the actual file information.
Private m_Rst As DAO.Recordset

' File details used for exporting/importing
Private m_Name As String
Private m_FileName As String
Private m_Extension As String
Private m_FileData() As Byte

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the shared image as a json file with file details, and a copy
'           : of the binary image file saved as an image.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()

    Dim strFile As String
    Dim strFolder As String
    Dim stm As ADODB.Stream
    Dim bteHeader As Byte
    Dim varHeader() As Byte
    Dim rst As Recordset2
    Dim rstAtc As Recordset2
    Dim strSql As String
    
    ' Save theme file
    strFile = IDbComponent_SourceFile & ".zip"
    strSql = "SELECT [Data] FROM MSysResources WHERE [Name]='" & m_Name & "' AND Extension='" & m_Extension & "'"
    Set m_Dbs = CurrentDb
    Set rst = m_Dbs.OpenRecordset(strSql, dbOpenSnapshot, dbOpenForwardOnly)
    
    ' If we get multiple records back we don't know which to use
    If rst.RecordCount > 1 Then Err.Raise 42, , "Multiple records in MSysResources table were found that matched name '" & m_Name & "' and extension '" & m_Extension & "' - Compact and repair database and try again."
    
    ' make sure parent folder exists before we try to save it
    VerifyPath FSO.GetParentFolderName(strFile)
    
    If Not rst.EOF Then
        Set rstAtc = rst!Data.Value
        If FSO.FileExists(strFile) Then FSO.DeleteFile strFile
        rstAtc!FileData.SaveToFile strFile
        rstAtc.Close
        Set rstAtc = Nothing
    End If
    rst.Close
    Set rst = Nothing
    
    ' Extract to folder and delete zip file.
    strFolder = IDbComponent_SourceFile
    If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder
    DoEvents ' Make sure the folder is deleted before we recreate it.
    ExtractFromZip strFile, IDbComponent_SourceFile, False
    ' Rather than holding up the export while we extract the file,
    ' use a cleanup sub to do this after the export.
    'FSO.DeleteFile IDbComponent_SourceFile & ".zip"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim rstResources As DAO.Recordset2
    Dim rstAttachment As DAO.Recordset2
    Dim fldFile As DAO.Field2
    Dim strZip As String
    Dim strThemeFile As String
    Dim strName As String
    Dim strSql As String
    
    ' Build zip file from theme folder
    strZip = strFile & ".zip"
    If FSO.FileExists(strZip) Then FSO.DeleteFile strZip
    DoEvents
    CreateZipFile strZip
    CopyFolderToZip strFile, strZip
    DoEvents

    ' Get theme name
    strName = FSO.GetBaseName(strZip)
    
    ' Create/edit record in resources table.
    VerifyResourcesTable
    strSql = "SELECT * FROM MSysResources WHERE [Type] = 'thmx' AND [Name]=""" & strName & """"
    Set rstResources = CurrentDb.OpenRecordset(strSql, dbOpenDynaset)
    With rstResources
        If .EOF Then
            ' No existing record found. Add a record
            .AddNew
            !Name = strName
            !Extension = "thmx"
            !Type = "thmx"
            Set rstAttachment = .Fields("Data").Value
        Else
            ' Found theme record with the same name.
            ' Remove the attached theme file.
            .Edit
            Set rstAttachment = .Fields("Data").Value
            If Not rstAttachment.EOF Then rstAttachment.Delete
        End If
        
        ' Upload theme file into OLE field
        strThemeFile = strFile & ".thmx"
        Name strZip As strThemeFile
        DoEvents
        With rstAttachment
            .AddNew
            Set fldFile = .Fields("FileData")
            fldFile.LoadFromFile strThemeFile
            .Update
            .Close
        End With
        
        ' Save and close record
        .Update
        .Close
    End With
    
    ' Remove zip file
    Kill strThemeFile
    
    ' Clear object (Important with DAO/ADO)
    Set rstAttachment = Nothing
    Set rstResources = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB() As Collection
    
    
    Dim cTheme As IDbComponent
    Dim rst As DAO.Recordset
    Dim strSql As String
    
    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        Set m_AllItems = New Collection
            
        ' This system table should exist, but just in case...
        If TableExists("MSysResources") Then

            Set m_Dbs = CurrentDb
            strSql = "SELECT * FROM MSysResources WHERE Type='thmx'"
            Set rst = m_Dbs.OpenRecordset(strSql, dbOpenSnapshot, dbOpenForwardOnly)
            With rst
                Do While Not .EOF
                    Set cTheme = New clsDbTheme
                    Set cTheme.DbObject = rst    ' Reference to OLE object recordset2
                    m_AllItems.Add cTheme, Nz(!Name)
                    .MoveNext
                Loop
                .Close
            End With
        End If
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems

End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyResourceTable
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Make sure the resources table exists, creating it if needed.
'---------------------------------------------------------------------------------------
'
Public Sub VerifyResourcesTable()

    Dim strName As String
    
    If Not TableExists("MSysResources") Then
        ' It would be nice to find a magical system command for this, but for now
        ' we can create it by creating a temporary form object.
        strName = CreateForm().Name
        ' Close without saving
        DoCmd.Close acForm, strName, acSaveNo
        ' Remove any potential default theme
        DoCmd.RunSQL "DELETE * FROM MSysResources WHERE [Type]='thmx'"
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Collection
    ' Get list of folders
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder, vbDirectory)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearOrphanedSourceFolders Me
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    IDbComponent_DateModified = 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceModified
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : The date/time the source object was modified. In most cases, this would
'           : be the date/time of the source file, but it some cases like SQL objects
'           : the date can be determined through other means, so this function
'           : allows either approach to be taken.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceModified() As Date
    '// TODO: Recursively identify the most recent file modified date.
    'If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = FileDateTime(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "themes"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "themes\"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = m_Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'           : In this case, we are building the name to include the info needed to
'           : recreate the record in the MSysResource table.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Name)
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count() As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB.Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbTheme
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    ' Not used
    Set IDbComponent_DbObject = Nothing
End Property


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_DbObject
' Author    : Adam Waller
' Date      : 5/11/2020
' Purpose   : Load in the class values from the recordset
'---------------------------------------------------------------------------------------
'
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)

    Dim fld2 As DAO.Field2
    Dim rst2 As DAO.Recordset2

    Set m_Rst = RHS

    ' Load in the object details.
    m_Name = m_Rst!Name
    m_Extension = m_Rst!Extension
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set fld2 = m_Rst!Data
    Set rst2 = fld2.Value
    m_FileName = rst2.Fields("FileName")
    m_FileData = rst2.Fields("FileData")

    ' Clear the object references
    Set rst2 = Nothing
    Set fld2 = Nothing
    Set m_Rst = Nothing

End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 5/13/2020
' Purpose   : Clear reference to database object.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set m_Dbs = Nothing
End Sub