VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'--------------------------------------------------------------------
' COPY BELOW SNIPPET TO CLASS, FORM OR MODULE AND REPLACE MODULE NAME
'--------------------------------------------------------------------
'Private logger As New logger
'Private Sub log(sMessage As String, Optional logLevel As logLevel = info)
'    If logger.IsInitialized = False Then
'        logger.Initialize ("<replace-with-module-name>")
'    End If
'    logger.log sMessage, logLevel
'End Sub

'-----------------
' PUBLIC TYPES
'-----------------
Public Enum logLevel
    Verbose = 1
    DebugLog = 2
    info = 3
    Warning = 4
    ErrorLog = 5
End Enum

'------------------
' PRIVATE VARIABLES
'------------------
Private m_sModuleName As String
Private m_isInitialized As Boolean
Private m_useSystemLog As Boolean
Private m_loggingEnabled As Boolean
Private m_logMessageAboveLevel As logLevel

'------------------
' PUBLIC PROPERTIES
'------------------
Public Property Get IsInitialized() As Boolean
    m_LogSystem ("IsInitialized")
    IsInitialized = m_isInitialized
End Property

'------------
' CONSTRUCTOR
'------------
Public Sub class_initialized()
    m_LogSystem ("Class_Initialized")
    m_isInitialized = False
End Sub

'---------------------------------
' PRIVATE PROCEDURES AND FUNCTIONS
'---------------------------------
Private Sub LogLoggerNotInitializedError()
    m_LogSystem ("LogLoggerNotInitializedError")
    m_sModuleName = "LOGGER"
    m_log "Logger is not initialized!", ErrorLog
End Sub

Private Function m_GetLogLevelString(level As logLevel) As String
    m_LogSystem ("m_GetLogLevelString")
    Select Case level
        Case 1
            m_GetLogLevelString = "VERBOSE"
        Case 2
            m_GetLogLevelString = "DEBUG"
        Case 3
            m_GetLogLevelString = "INFO"
        Case 4
            m_GetLogLevelString = "WARNING"
        Case 5
            m_GetLogLevelString = "ERROR"
    End Select
End Function

Private Sub m_log(sLogMessage As String, Optional level As logLevel = logLevel.info)
    m_LogSystem ("m_log")
    Dim writer As TextStream, sMessage As String, sLogLevel As String, today As String
    
    If IsNull(sLogMessage) Then
        sLogMessage = "NULL"
    
    End If
    
    today = m_GetCurrentDate
    Dim logFile As String
    logFile = "mainchef-" & today & ".log"
    
    sLogLevel = m_GetLogLevelString(level)
    sMessage = Now & "|" & sLogLevel & "|" & m_sModuleName & "|" & sLogMessage
    
    With m_GetLogFileWriter(logFile)
        .WriteLine (sMessage)
    End With
    
    Debug.Print sMessage
End Sub

Private Function m_GetCurrentDateAndTime() As String
    m_LogSystem ("m_GetCurrentDateAndTime")
    m_GetCurrentDateAndTime = m_GetCurrentDate & "T" & Time
End Function

Private Function m_GetCurrentDate() As String
    m_LogSystem ("m_GetCurrentDate")
    On Error Resume Next
    Dim today As String, dateSplit, sSplitChar As String: sSplitChar = "-"
    m_LogSystem ("Date is: " & Date & " - splitting on '" & sSplitChar & "'")
    dateSplit = Split(Date, sSplitChar)
    today = dateSplit(2) & dateSplit(1) & dateSplit(0)
    
    m_GetCurrentDate = today
    
    If Err <> 0 Then
        m_LogSystem ("Error happened in m_GetCurrentDate")
    End If
    
End Function

Private Function m_GetLogFileWriter(sLogFileName As String)
    Dim fs As Scripting.FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim logDir As String
    logDir = CurrentProject.Path & "\logs"
    
    If Not fs.FolderExists(logDir) Then
        fs.CreateFolder (logDir)
    End If

    Dim sLogFilePath As String
    sLogFilePath = logDir & "\" & sLogFileName
    Set m_GetLogFileWriter = fs.OpenTextFile(sLogFilePath, 8, True)
    
End Function

Private Sub m_LogSystem(sMessage As String)
    If m_useSystemLog Then
        With m_GetLogFileWriter("system.log")
            .WriteLine (Now & ": " & sMessage)
        End With
    End If
End Sub


'--------------------------------
' PUBLIC PROCEDURES AND FUNCTIONS
'--------------------------------
Public Function Initialize(sModuleName As String, Optional bDisableSystemLog As Boolean = False)
    m_LogSystem ("Initialize")
    m_sModuleName = sModuleName
    m_isInitialized = True
    
    If bDisableSystemLog = False Then
        m_useSystemLog = False ' SettingsProvider.GetSetting("systemLogging")
    Else
        m_useSystemLog = False
    End If
    
    m_loggingEnabled = True '(SettingsProvider.GetSetting("disableAppLogging") = False)
    m_logMessageAboveLevel = 1 ' SettingsProvider.GetLogLevel 'IIf((SettingsProvider.GetSetting("logLevel") = False), logLevel.info, SettingsProvider.GetSetting("logLevel"))
    log "Logger initialized for module '" & sModuleName & "'"
End Function


Public Sub log(logMessage As String, Optional logLevel As logLevel = info)
    m_LogSystem ("log")
    If IsInitialized = False Then
        LogLoggerNotInitializedError
        Err.Raise 1, "Logger is not initialized!"
    End If
    
    If m_loggingEnabled And logLevel >= m_logMessageAboveLevel Then
        m_log logMessage, logLevel
    End If
End Sub

