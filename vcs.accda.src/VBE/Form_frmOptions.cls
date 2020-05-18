VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : NOTE - Options are dynamically loaded to the controls, and then saved
'           : back to the class (and subsequently project) when the user clicks the
'           : save and close button. Most options will not require event coding since
'           : they are mapped to their appropriate option using the control name.
'           : I.e. chkUseFastSave = cOptions.UseFastSave
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private cOptions As clsOptions

' Use a collection to represent tables, since the user can filter which ones they
' want to see while updating items in the list. We don't want to save till the very
' end when we write this information back to the json options file.
Private m_colTables As Collection

Private Enum eTableCol
    etcName = 0
    etcType = 1
    etcHidden = 2
    etcSystem = 3
    etcOther = 4
End Enum

Private Enum eMapAction
    emaClassToForm
    emaFormToClass
End Enum


'---------------------------------------------------------------------------------------
' Procedure : chkTableShowHidden_Click
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Update the list of tables
'---------------------------------------------------------------------------------------
'
Private Sub chkTableShowHidden_Click()
    RefreshTableDisplay
End Sub
Private Sub chkTableShowOther_Click()
    RefreshTableDisplay
End Sub
Private Sub chkTableShowSystem_Click()
    RefreshTableDisplay
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdCancel_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Cancel (close) the form.
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdClearDefaults_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Clear the default and project settings.
'---------------------------------------------------------------------------------------
'
Private Sub cmdClearDefaults_Click()
    Dim cDefaults As clsOptions
    Set cDefaults = New clsOptions
    cDefaults.SaveOptionsAsDefault
    If MsgBox2("Apply to this Project?", _
        "The default options have been reset. Would you like these applied to this project as well?", _
        "Click 'Yes' to apply the default options, or 'No' to leave the current options.", _
        vbQuestion + vbYesNo, "Version Control System") = VbMsgBoxResult.vbYes Then
        cmdRestoreDefaults_Click
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadTableList
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Load the list of tables in the current database, and merge in the list
'           : of tables where we are opting to save data.
'---------------------------------------------------------------------------------------
'
Private Function LoadTableList()

    Dim tbl As AccessObject
    Dim blnHidden As Boolean
    Dim blnSystem As Boolean
    Dim blnShow As Boolean
    Dim blnOther As Boolean
    Dim intIndex As Integer
    Dim dTable As Dictionary
    Dim varTable As Variant
    Dim varKey As Variant
    Dim intFormat As eTableDataExportFormat
    Dim strFormat As String
    Dim strName As String
    
    ' Reset list of tables
    Set m_colTables = New Collection
    
    ' Get list of tables
    For Each tbl In CurrentData.AllTables
        ' Read table attributes
        blnHidden = Application.GetHiddenAttribute(acTable, tbl.Name)
        blnSystem = (tbl.Attributes And dbSystemObject)
        blnOther = False    ' Other represents tables not in this database.
        ' Add array record to represent table.
        m_colTables.Add Array(tbl.Name, vbNullString, blnHidden, blnSystem, blnOther)
    Next tbl
    
    ' Add in the list of saved tables, adding into the sorted location
    If Not cOptions.TablesToExportData Is Nothing Then
        ' Loop through each table in the saved table list
        For Each varKey In cOptions.TablesToExportData.Keys
            strFormat = CStr(cOptions.TablesToExportData.Item(varKey)("Format"))
            intFormat = TableExportFormatKeyToEnum(strFormat)
            strName = CStr(varKey)
            AddUpdateTableInList strName, strFormat, False, False, True
        Next varKey
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddUpdateTableInList
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Updates the a table in the collection, adding it if it does not exist.
'---------------------------------------------------------------------------------------
'
Private Sub AddUpdateTableInList(strName As String, strFormat As String, blnHidden As Boolean, blnSystem As Boolean, blnOther As Boolean)

    Dim intIndex As Integer
    Dim varTable As Variant
    
    ' Now loop through the collection of tables, inserting it into the alphabetical position.
    For intIndex = 1 To m_colTables.Count
        ' First, see if the table name was already loaded into the collection.
        varTable = m_colTables(intIndex)
        If varTable(etcName) = strName Then
            ' Found matching table name. Update the existing record
            SetTableExportType strName, strFormat
            Exit Sub
        Else
            ' Check alphabetical position.
            If (varTable(etcName) > strName) Then
                ' Our saved table should be inserted before this record.
                m_colTables.Add Array(strName, strFormat, False, False, True), strName, intIndex
                Exit Sub
            End If
        End If
    Next intIndex
    
    ' The table might need to be added to the very end of the list.
    m_colTables.Add Array(strName, strFormat, False, False, True), strName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TableExportFormatKeyToEnum
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Translate the table export format key to the corresponding enum value.
'---------------------------------------------------------------------------------------
'
Private Function TableExportFormatKeyToEnum(strKey As String) As eTableDataExportFormat
    Dim intFormat As eTableDataExportFormat
    Dim strName As String
    For intFormat = etdNoData To eTableDataExportFormat.[_last]
        strName = cOptions.GetTableExportFormatName(intFormat)
        If strName = strKey Then
            TableExportFormatKeyToEnum = intFormat
            Exit For
        End If
    Next intFormat
End Function


'---------------------------------------------------------------------------------------
' Procedure : TableExportFormatNameToEnum
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Translate the friendly name of the table export format to the enum.
'---------------------------------------------------------------------------------------
'
Private Function TableExportFormatNameToEnum(strKey As String) As eTableDataExportFormat
    Dim intFormat As eTableDataExportFormat
    Dim strName As String
    For intFormat = etdNoData To eTableDataExportFormat.[_last]
        strName = cOptions.GetTableExportFormatName(intFormat)
        If strName = strKey Then
            TableExportFormatNameToEnum = intFormat
            Exit For
        End If
    Next intFormat
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetTableExportType
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Set the export type for the table
'---------------------------------------------------------------------------------------
'
Private Function SetTableExportType(strTable As String, strFormat As String)

    Dim varTable As Variant
    Dim varNew As Variant
    Dim strType As String
    Dim intIndex As Integer
    
    For Each varTable In m_colTables
        intIndex = intIndex + 1
        If varTable(etcName) = strTable Then
            ' Create new record
            strType = Replace(strFormat, "No Data", vbNullString)
            varNew = Array(strTable, strType, varTable(etcHidden), varTable(etcSystem), varTable(etcOther))
            ' Replace with new record
            m_colTables.Remove intIndex
            If intIndex = 1 Then
                m_colTables.Add varNew, strTable, intIndex
            Else
                m_colTables.Add varNew, strTable, , intIndex - 1
            End If
            Exit For
        End If
    Next varTable

End Function


'---------------------------------------------------------------------------------------
' Procedure : RefreshTableDisplay
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : The master list of tables is stored in the m_colTables
'---------------------------------------------------------------------------------------
'
Private Function RefreshTableDisplay()
    
    Dim varTable As Variant
    Dim varValue As Variant
    Dim blnShow As Boolean
    
    ' Update the list box of table names from the collection.
    With Me.lstTables

        ' Add header row
        .RowSource = ""
        .AddItem "Table Name;Save Data"
        
        ' Get list of tables
        If Not m_colTables Is Nothing Then
            For Each varTable In m_colTables
                ' Show all, unless hidden by options
                blnShow = Not (varTable(etcHidden) Or varTable(etcSystem))
                ' Evaluate hidden attribute
                If Me.chkTableShowHidden And varTable(etcHidden) Then blnShow = True
                ' Check system attribute
                If Me.chkTableShowSystem And varTable(etcSystem) Then blnShow = True
                ' Check for other tables
                If Me.chkTableShowOther And varTable(etcOther) Then blnShow = True
                ' Add to list
                If blnShow Then .AddItem varTable(etcName) & ";" & varTable(etcType)
            Next varTable
        End If
    End With
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : cmdRestoreDefaults_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Restore the default options to this project.
'---------------------------------------------------------------------------------------
'
Private Sub cmdRestoreDefaults_Click()
    cOptions.LoadDefaultOptions
    MapControlsToOptions emaClassToForm
    MsgBox2 "Default Options Restored", _
        "The system default options have been restored to this project.", _
        "Be sure to click Save and Close when you are finished making changes.", _
        vbInformation, "Version Control System"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSaveAndClose_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save and close the options dialog
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveAndClose_Click()
    MapControlsToOptions emaFormToClass
    cOptions.SaveOptionsForProject
    DoCmd.Close acForm, Me.Name
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSaveAsDefault_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Save current options as default for projects.
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveAsDefault_Click()
    cOptions.SaveOptionsAsDefault
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdUpdateTableData_Click
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Update the table in the list
'---------------------------------------------------------------------------------------
'
Private Sub cmdUpdateTableData_Click()
    If Nz(lstTables, 0) > 0 Then
        SetTableExportType lstTables, Nz(cboTableDataSaveType)
        RefreshTableDisplay
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Load options for this project
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim intFormat As eTableDataExportFormat
    
    Set cOptions = New clsOptions
    cOptions.LoadProjectOptions
    MapControlsToOptions emaClassToForm
    RefreshTableDisplay
    
    ' Load list of table data export formats
    cboTableDataSaveType.AddItem "No Data"
    For intFormat = 1 To eTableDataExportFormat.[_last]
        cboTableDataSaveType.AddItem cOptions.GetTableExportFormatName(intFormat)
    Next intFormat
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MapControlsToOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Map the form controls to the options, performing the specified action.
'---------------------------------------------------------------------------------------
'
Private Function MapControlsToOptions(eAction As eMapAction)

    Dim pge As Access.Page
    Dim ctl As Control
    Dim strKey As String
    Dim varItem As Variant
    Dim dTables As Scripting.Dictionary
    Dim dTable As Scripting.Dictionary
    
    ' Loop through each page
    For Each pge In tabOptions.Pages
        For Each ctl In pge.Controls
            Select Case TypeName(ctl)
                Case "CheckBox", "TextBox"
                    Select Case ctl.Name
                        Case "chkTableShowHidden", "chkTableShowSystem", "chkTableShowOther", "txtTableName"
                            ' Skip these exceptions.
                        Case Else
                            ' Get option name from control name following prefix.
                            strKey = Mid(ctl.Name, 4)
                            If eAction = emaClassToForm Then
                                ctl = CallByName(cOptions, strKey, VbGet)
                            ElseIf eAction = emaFormToClass Then
                                CallByName cOptions, strKey, VbLet, Nz(ctl.Value)
                            End If
                    End Select
                Case "ListBox"
                    ' Special handling for lists
                Case Else
                    ' Ignore other controls
            End Select
        Next ctl
    Next

    ' Table list for selecting tables to save data.
    If eAction = emaClassToForm Then
        LoadTableList
    ElseIf eAction = emaFormToClass Then
        ' Save list of tables to export data
        Set dTables = New Scripting.Dictionary
        For Each varItem In m_colTables
            If varItem(etcType) <> vbNullString Then
                Set dTable = New Scripting.Dictionary
                dTable("Format") = varItem(etcType)
                'dTable("ObjectType") = "Table"  ' Could add queries later
                dTables.Add varItem(etcName), dTable
            End If
        Next varItem
        Set cOptions.TablesToExportData = dTables
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : lblAddOtherTable_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Select another table not used in the current database
'---------------------------------------------------------------------------------------
'
Private Sub lblAddOtherTable_Click()
    Dim strTable As String
    If Nz(cboTableDataSaveType, "No Data") = "No Data" Then
        MsgBox2 "Please select format", "Select the format to save table data before adding the table to the list.", , vbInformation
    Else
        strTable = InputBox("Enter the name of a table that doesn't exist in this database, " & _
            "but from which you would like to export data if it is found.", "Enter Table Name")
        If strTable <> vbNullString Then
            AddUpdateTableInList strTable, Nz(cboTableDataSaveType), False, False, True
            chkTableShowOther = True
            RefreshTableDisplay
        End If
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lstTables_Click
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Update the data type selection to match the table row
'---------------------------------------------------------------------------------------
'
Private Sub lstTables_Click()
    With lstTables
        If Nz(.Value) = vbNullString Then
            cboTableDataSaveType = vbNullString
        Else
            cboTableDataSaveType = .Column(1, .ListIndex + 1)
        End If
        If Nz(cboTableDataSaveType) = vbNullString Then cboTableDataSaveType = "No Data"
        txtTableName = Nz(.Value)
    End With
End Sub

