Attribute VB_Name = "modAddIn"
Option Compare Database
Option Explicit
Option Private Module

Public Enum eReleaseType
    Major_Vxx = 0
    Minor_xVx = 1
    Build_xxV = 2
End Enum

' Used to determine if Access is running as administrator. (Required for installing the add-in)
Private Declare PtrSafe Function IsUserAnAdmin Lib "shell32" () As Long

' Used to relaunch Access as an administrator to install the addin.
#If Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
#End If

Private Const SW_SHOWNORMAL = 1
Private Logger As New Logger
Private Sub log(sMessage As String, Optional logLevel As logLevel = info)
    If Logger.IsInitialized = False Then
        Logger.Initialize ("modAddIn")
    End If
    Logger.log sMessage, logLevel
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemLaunch
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemLaunch()
    DoCmd.OpenForm "frmMain"
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemExport
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Open main form and start export immediately. (Save users a click)
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemExport()
    DoCmd.OpenForm "frmMain"
    DoEvents
    Form_frmMain.cmdExport_Click
End Function


'---------------------------------------------------------------------------------------
' Procedure : AutoRun
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : This code runs when the add-in file is opened directly. It provides the
'           : user an easy way to update the add-in on their system.
'---------------------------------------------------------------------------------------
'
Public Function AutoRun()
    log "AutoRun"
    ' If we are running from the addin location, we might be trying to register it.
    If CodeProject.FullName = GetAddinFileName Then
        log "Might be trying to register addin"
        
        ' See if the user has admin privileges
        If IsUserAnAdmin = 1 Then
            log "User is admin"
            ' Create the menu items
            ' NOTE: Be sure to keep these consistent with the USysRegInfo table
            ' so the user can uninstall the add-in later if desired.
            log "Register menu items"
            'RegisterMenuItem "VCS &Version Control", "=AddInMenuItemLaunch()"
            'RegisterMenuItem "VCS &Export All Source", "=AddInMenuItemExport()"
            InstalledVersion = AppVersion
            
            ' Give success message and quit Access
            If IsAlreadyInstalled Then
                MsgBox2 "Success!", "Version Control System has now been installed.", _
                    "You may begin using this tool after reopening Microsoft Access", vbInformation, "Version Control Add-in"
                'DoCmd.Quit
            End If
        Else
            ' User does not have admin priviledges. Shouldn't normally be opening the add-in directly.
            ' Don't do anything special here. Just let them browse around in the file.
        End If
    Else
        log "This DB is not running from the addin directory"
        ' Could be running it from another location, such as after downloading
        ' and updated version of the addin. In that case, we are either trying
        ' to install it for the first time, or trying to upgrade it.
        If IsAlreadyInstalled Then
            log "Version Control is already installed"
            If InstalledVersion <> AppVersion Then
                log "The installed version is different from AppVersion"
                If MsgBox2("Upgrade Version Control?", _
                    "Would you like to upgrade to version " & AppVersion & "?", _
                    "Click 'Yes' to continue or 'No' to cancel.", vbQuestion + vbYesNo, "Version Control Add-in") = vbYes Then
                    log "User want to upgrade Version Control"
                    If InstallVCSAddin Then
                        log "Install VCS Addin succeeded"
                        MsgBox2 "Success!", "Version Control System add-in has been updated to " & AppVersion & ".", _
                            "Please restart any open instances of Microsoft Access before using the add-in.", vbInformation, "Version Control Add-in"
                        log "Quit Access"
                        DoCmd.Quit
                    End If
                End If
            End If
        Else
            log "Version Control is not installed"
            ' Not yet installed. Offer to install.
            If MsgBox2("Install Version Control?", _
                "Would you like to install version " & AppVersion & "?", _
                "Click 'Yes' to continue or 'No' to cancel.", vbQuestion + vbYesNo, "Version Control Add-in") = vbYes Then
                log "User wants to install"
                
                If InstallVCSAddin Then
                    log "Install VCS addin suceeded - relaunch as admin"
                    RelaunchAsAdmin
                End If
                log "Quit access"
                DoCmd.Quit
            End If
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : InstallVCSAddin
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Installs/updates the add-in for the current user.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Private Function InstallVCSAddin()
    log "InstallVCSAddin"
    Dim strSource As String
    Dim strDest As String

    Dim blnExists As Boolean
    
    strSource = CodeProject.FullName
    strDest = GetAddinFileName
    
    ' We can't replace a file with itself.  :-)
    If strSource = strDest Then
        log "Source is same as destination. Quit Access"
        Exit Function
    End If
    ' Copy the file, overwriting any existing file.
    ' Requires FSO to copy open database files. (VBA.FileCopy give a permission denied error.)
    On Error Resume Next
    log "Copy " & strSource & " to " & strDest
    FSO.CopyFile strSource, strDest, True
    If Err Then
        log "Error happened: " & Err.Description
        Err.Clear
    Else
        log "No error in copy"
        ' Update installed version number
        InstalledVersion = AppVersion
        ' Return success
        log "Install Version Control succeeded"
        InstallVCSAddin = True
    End If
    On Error GoTo 0
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddinFileName
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : This is where the add-in would be installed.
'---------------------------------------------------------------------------------------
'
Private Function GetAddinFileName() As String
    log "GetAddinFileName"
    Dim fileName As String: fileName = Environ("AppData") & "\Microsoft\AddIns\" & CodeProject.Name
    log "Addin filename: " & fileName
    GetAddinFileName = fileName
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsAlreadyInstalled
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Returns true if the addin is already installed.
'---------------------------------------------------------------------------------------
'
Private Function IsAlreadyInstalled() As Boolean
    log "IsAlreadyInstalled"
    Dim installed As Boolean
    Dim strPath As String
    Dim oShell As IWshRuntimeLibrary.WshShell
    Dim strTest As String
    
    ' Check for registry key of installed version
    If InstalledVersion <> vbNullString Then
        log "Installed version is different from null"
        
        ' Check for addin file
        If Dir(GetAddinFileName) = CodeProject.Name Then
            log "Addin filename is same as CodeProject.Name"
            
            ' Check HKLM registry key
            Set oShell = New IWshRuntimeLibrary.WshShell
            strPath = GetAddinRegPath & "Version Control\Library"
            log "RegPath: " & strPath
            On Error Resume Next
            ' We should have a value here if the install ran in the past.
            strTest = oShell.RegRead(strPath)
            
            If Err Then
                log "Error happended: " & Err.Description, ErrorLog
                log "Value from regpath: " & strTest
                Err.Clear
            End If
            On Error GoTo 0
            Set oShell = Nothing
        
            ' Return our determination
            installed = (strTest <> vbNullString)
        End If
    End If
    
    log "Is installed: " & installed
    IsAlreadyInstalled = installed
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddinRegPath
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Return the registry path to the addin menu items
'---------------------------------------------------------------------------------------
'
Private Function GetAddinRegPath() As String
    Dim regPath As String: regPath = "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\" & Application.Version & "\Access\Menu Add-Ins\"
    log "GetAddinRegPath: " & regPath
    GetAddinRegPath = regPath
End Function


'---------------------------------------------------------------------------------------
' Procedure : RegisterMenuItem
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Add the menu item through the registry (HKLM, requires admin)
'---------------------------------------------------------------------------------------
'
Private Function RegisterMenuItem(strName, Optional strFunction As String = "=LaunchMe()")
    log "RegisterMenuItem"
    Dim oShell As IWshRuntimeLibrary.WshShell
    Dim strPath As String
    
    Set oShell = New IWshRuntimeLibrary.WshShell
    
    ' We need to create/update three registry keys for each item.
    strPath = GetAddinRegPath & strName & "\"
    With oShell
        log "Write to registry: " & strPath & "Expression"
        .RegWrite strPath & "Expression", strFunction, "REG_SZ"
        log "Write to registry: " & strPath & "Library"
        .RegWrite strPath & "Library", GetAddinFileName, "REG_SZ"
        log "Write to registry: " & strPath & "Version"
        .RegWrite strPath & "Version", 3, "REG_DWORD"
    End With
    Set oShell = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RelaunchAsAdmin
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Launch the addin file with admin privileges so the user can register it.
'---------------------------------------------------------------------------------------
'
Private Sub RelaunchAsAdmin()
    log "RelaunchAsAdmin"
    ShellExecute 0, "runas", SysCmd(acSysCmdAccessDir) & "\msaccess.exe", """" & GetAddinFileName & """", vbNullString, SW_SHOWNORMAL
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Deploy
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Increments the build version and updates the project description.
'           : This can be run from the debug window when making updates to the project.
'           : (More significant updates to the version number can be made using the
'           :  `AppVersion` property defined below.)
'---------------------------------------------------------------------------------------
'
Public Sub Deploy(Optional ReleaseType As eReleaseType = Build_xxV)
    log "Deploy"
    Const cstrSpacer As String = "--------------------------------------------------------------"
        
    ' Make sure we don't run ths function while it is loaded in another project.
    If CodeProject.FullName <> CurrentProject.FullName Then
        Debug.Print "This can only be run from a top-level project."
        Debug.Print "Please open " & CodeProject.FullName & " and try again."
        Exit Sub
    End If
    
    ' Increment build number
    IncrementBuildVersion ReleaseType
    
    ' List project and new build number
    Debug.Print cstrSpacer
    
    ' Update project description
    VBE.ActiveVBProject.Description = "Version " & AppVersion & " deployed on " & Date
    Debug.Print " ~ " & VBE.ActiveVBProject.Name & " ~ Version " & AppVersion
    Debug.Print cstrSpacer
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IncrementBuildVersion
' Author    : Adam Waller
' Date      : 1/6/2017
' Purpose   : Increments the build version (1.0.12)
'---------------------------------------------------------------------------------------
'
Public Sub IncrementBuildVersion(ReleaseType As eReleaseType)
    log "IncrementBuildVersion"
    Dim varParts As Variant
    
    varParts = Split(AppVersion, ".")
    
    If UBound(varParts) <> 2 Then
        Debug.Print "Unexpected version format"
        Stop
    End If
    
    If Not IsNumeric(varParts(ReleaseType)) Then
        Debug.Print "Expecting numeric value"
        Stop
    Else
        varParts(ReleaseType) = varParts(ReleaseType) + 1
        AppVersion = Join(varParts, ".")
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Get the version from the database property.
'---------------------------------------------------------------------------------------
'
Public Property Get AppVersion() As String
    log "AppVersion"
    Dim strVersion As String
    strVersion = GetDBProperty("AppVersion")
    If strVersion = "" Then strVersion = "1.0.0"
    log "App Version: " & strVersion
    AppVersion = strVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Set version property in current database.
'---------------------------------------------------------------------------------------
'
Public Property Let AppVersion(strVersion As String)
    log "Let AppVersion: " & strVersion
    SetDBProperty "AppVersion", strVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Get a database property
'---------------------------------------------------------------------------------------
'
Public Function GetDBProperty(strName As String) As Variant
    log "GetDBProperty: " & strName
    Dim prp As DAO.Property
    
    For Each prp In CodeDb.Properties
        If prp.Name = strName Then
            GetDBProperty = prp.Value
            Exit For
        End If
    Next prp
    
    Set prp = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Set a database property
'---------------------------------------------------------------------------------------
'
Public Sub SetDBProperty(strName As String, varValue, Optional prpType = DB_TEXT)
    log "SetDBProperty: " & strName & " = " & varValue
    
    Dim prp As DAO.Property
    Dim blnFound As Boolean
    Dim dbs As DAO.Database
    
    Set dbs = CodeDb
    
    For Each prp In dbs.Properties
        If prp.Name = strName Then
            blnFound = True
            ' Skip set on matching value
            If prp.Value = varValue Then
                Set dbs = Nothing
                Exit Sub
            End If
            Exit For
        End If
    Next prp
    
    On Error Resume Next
    If blnFound Then
        dbs.Properties(strName).Value = varValue
    Else
        Set prp = dbs.CreateProperty(strName, DB_TEXT, varValue)
        dbs.Properties.Append prp
    End If
    If Err Then Err.Clear
    On Error GoTo 0
    
    Set dbs = Nothing
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : InstalledVersion
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Returns the installed version of the add-in from the registry.
'           : (We are saving this in the user hive, since it requires admin rights
'           :  to change the keys actually used by Access to register the add-in)
'---------------------------------------------------------------------------------------
'
Private Property Let InstalledVersion(strVersion As String)
    log "InstalledVersion: " & strVersion
    SaveSetting VBE.ActiveVBProject.Name, "Add-in", "Installed Version", strVersion
End Property
Private Property Get InstalledVersion() As String
    log "InstalledVersion"
    InstalledVersion = GetSetting(VBE.ActiveVBProject.Name, "Add-in", "Installed Version", vbNullString)
End Property
