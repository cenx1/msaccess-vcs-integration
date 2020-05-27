VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' This color scheme can be changed, I just wanted something more aesthetically
' pleasing than the default wizards and forms.
' Color scheme: https://coolors.co/383f51-e0e0e6-ffffff-ef8354-d3d7ef

'---------------------------------------------------------------------------------------
' Procedure : cmdClose_Click
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Close the form
'---------------------------------------------------------------------------------------
'
Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdExport_Click
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Export source code from current database
'---------------------------------------------------------------------------------------
'
Public Sub cmdExport_Click()
    
    cmdClose.SetFocus
    HideActionButtons
    DoEvents
    txtLog.Visible = True
    txtLog.SetFocus
    
    ' Show the status
    SetStatusText "Running...", "Exporting source code", "A summary of the export progress can be seen on this screen, and additional details are included in the log file."
    
    ' Export the source code
    modImportExport.ExportAllSource
    
    SetStatusText "Finished", "Export Complete", "Additional details can be found in the project export log file.<br><br>You may now close this window."
    cmdClose.SetFocus
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HideActionButtons
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Hide the action buttons when running a command.
'---------------------------------------------------------------------------------------
'
Private Sub HideActionButtons()
    cmdExport.Visible = False
    cmdImport.Visible = False
    cmdOptions.Visible = False
    cmdHelp.Visible = False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdHelp_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Go to the GitHub documentation page.
'---------------------------------------------------------------------------------------
'
Private Sub cmdHelp_Click()
    Application.FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration/wiki/Documentation"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdOptions_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Open options dialog (for this project)
'---------------------------------------------------------------------------------------
'
Private Sub cmdOptions_Click()
    DoCmd.OpenForm "frmOptions"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Setting the control source causes delayed display. This way the display
'           : is instant when the form is opened.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    lblVersion.Caption = "Version " & GetVCSVersion()
    SetStatusText "Choose Action", "What would you like to do?", _
        "<strong><em>Export</em></strong> source to generate source files from the current database." & _
        "<br><br><strong><em>Import</em></strong> source files to rebuild this database from source."
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetStatusText
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Update the status text to direct the user to the next task, or show the
'           : status of the current operation.
'---------------------------------------------------------------------------------------
'
Private Sub SetStatusText(strHeading As String, strSubHeading As String, strDescriptionHtml As String)
    lblHeading.Caption = strHeading
    lblSubheading.Caption = strSubHeading
    txtDescription.Value = strDescriptionHtml
End Sub
