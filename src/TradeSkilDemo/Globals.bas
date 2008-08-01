Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================

Public Const AppName                       As String = "TradeSkil Demo Edition"

Public Const IncreasedValueColor            As Long = &HB7E43
Public Const DecreasedValueColor            As Long = &H4444EB

' command line switch indicating which configuration to load
' when the programs starts (if not specified, the default configuration
' is loaded)
Public Const SwitchConfig                   As String = "config"

' command line switch specifying the log filename
Public Const SwitchLogFilename              As String = "log"

' command line switch specifying the loglevel
Public Const SwitchLogLevel                 As String = "loglevel"

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mStudyPickerForm                    As fStudyPicker

'@================================================================================
' Properties
'@================================================================================

Public Property Get gCommandLineParser() As CommandLineParser
Static clp As CommandLineParser
If clp Is Nothing Then Set clp = CreateCommandLineParser(Command)
Set gCommandLineParser = clp
End Property

Public Property Get gLogFileName() As String
Static logFileName As String
If logFileName = "" Then
    If gCommandLineParser.Switch(SwitchLogFilename) Then logFileName = gCommandLineParser.SwitchValue(SwitchLogFilename)

    If logFileName = "" Then
        logFileName = GetSpecialFolderPath(FolderIdLocalAppdata) & _
                                            "\TradeWright\" & _
                                            AppName & _
                                            "\v" & _
                                            App.Major & "." & App.Minor & _
                                            "\log.txt"
    End If
End If
gLogFileName = logFileName
End Property

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("log")
Set gLogger = lLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gShowStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal title As String)
If mStudyPickerForm Is Nothing Then Set mStudyPickerForm = New fStudyPicker
mStudyPickerForm.initialise chartMgr, title
mStudyPickerForm.Show vbModeless
End Sub

Public Sub gSyncStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal title As String)
If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise chartMgr, title
End Sub

Public Sub gUnsyncStudyPicker()
If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise Nothing, "Study picker"
End Sub


