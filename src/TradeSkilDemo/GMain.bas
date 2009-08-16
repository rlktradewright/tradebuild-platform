Attribute VB_Name = "GMain"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                            As String = "TradeSkilDemo26"
Private Const ModuleName                            As String = "GMain"

Public Const AppName                                As String = "TradeSkil Demo Edition"

Public Const ConfigFileVersion                      As String = "1.1"

Public Const ConfigSectionChart                     As String = "Chart"
Public Const ConfigSectionCharts                    As String = "Charts"
Public Const ConfigSectionConfigEditor              As String = "ConfigEditor"
Public Const ConfigSectionDefaultStudyConfigs       As String = "DefaultStudyConfigs"
Public Const ConfigSectionMainForm                  As String = "MainForm"
Public Const ConfigSectionMultiChart                As String = "MultiChart"
Public Const ConfigSectionOrderTicket               As String = "OrderTicket"
Public Const ConfigSectionTickerGrid                As String = "TickerGrid"

Public Const ConfigSettingHeight                    As String = ".Height"
Public Const ConfigSettingLeft                      As String = ".Left"
Public Const ConfigSettingTop                       As String = ".Top"
Public Const ConfigSettingWidth                     As String = ".Width"
Public Const ConfigSettingWindowstate               As String = ".Windowstate"

Public Const ConfigSettingConfigEditorLeft          As String = ConfigSectionConfigEditor & ConfigSettingLeft
Public Const ConfigSettingConfigEditorTop           As String = ConfigSectionConfigEditor & ConfigSettingTop

Public Const ConfigSettingMainFormControlsHidden    As String = ConfigSectionMainForm & ".ControlsHidden"
Public Const ConfigSettingMainFormFeaturesHidden    As String = ConfigSectionMainForm & ".FeaturesHidden"
Public Const ConfigSettingMainFormHeight            As String = ConfigSectionMainForm & ConfigSettingHeight
Public Const ConfigSettingMainFormLeft              As String = ConfigSectionMainForm & ConfigSettingLeft
Public Const ConfigSettingMainFormTop               As String = ConfigSectionMainForm & ConfigSettingTop
Public Const ConfigSettingMainFormWidth             As String = ConfigSectionMainForm & ConfigSettingWidth
Public Const ConfigSettingMainFormWindowstate       As String = ConfigSectionMainForm & ConfigSettingWindowstate

Public Const ConfigSettingOrderTicketLeft           As String = ConfigSectionOrderTicket & ConfigSettingLeft
Public Const ConfigSettingOrderTicketTop            As String = ConfigSectionOrderTicket & ConfigSettingTop

Private Const DefaultAppInstanceConfigName          As String = "Default Config"

Public Const LB_SETHORZEXTENT                       As Long = &H194

' the SSTAB control subtracts this amount from the Left property of controls
' that are not on the active tab to ensure they aren't visible
Public Const SSTabInactiveControlAdjustment         As Long = 75000

' command line switch indicating which configuration to load
' when the programs starts (if not specified, the default configuration
' is loaded)
Public Const SwitchConfig                           As String = "config"

' command line switch specifying the log filename
Public Const SwitchLogFilename                      As String = "log"

' command line switch specifying the loglevel
Public Const SwitchLogLevel                         As String = "loglevel"

Public Const WindowStateMaximized                   As String = "Maximized"
Public Const WindowStateMinimized                   As String = "Minimized"
Public Const WindowStateNormal                      As String = "Normal"

'@================================================================================
' Member variables
'@================================================================================

Public gConfigFile                                  As ConfigurationFile
Public gAppInstanceConfig                           As ConfigurationSection

Private mStudyPickerForm                            As fStudyPicker
Private mMainForm                                   As fTradeSkilDemo

Private mEditConfig                                 As Boolean

Private mListener                                   As LogListener

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get gAppTitle() As String
gAppTitle = AppName & _
                " v" & _
                App.Major & "." & App.Minor
End Property

Public Property Get gCommandLineParser() As CommandLineParser
Static clp As CommandLineParser
If clp Is Nothing Then Set clp = CreateCommandLineParser(Command)
Set gCommandLineParser = clp
End Property

Public Property Get gAppSettingsFolder() As String
gAppSettingsFolder = GetSpecialFolderPath(FolderIdLocalAppdata) & _
                    "\TradeWright\" & _
                    gAppTitle
End Property

Public Property Get gLogFileName() As String
Static logFileName As String
If logFileName = "" Then
    If gCommandLineParser.Switch(SwitchLogFilename) Then logFileName = gCommandLineParser.SwitchValue(SwitchLogFilename)

    If logFileName = "" Then
        logFileName = gAppSettingsFolder & "\log.txt"
    End If
End If
gLogFileName = logFileName
End Property

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("log")
Set gLogger = lLogger
End Property

Public Property Get gMainForm() As fTradeSkilDemo
Set gMainForm = mMainForm
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gHandleFatalError(ByVal errNum As Long, _
                            ByVal Description As String, _
                            ByVal Source As String)
gLogger.Log LogLevelSevere, _
        "Error number: " & errNum & vbCrLf & _
        "Description: " & Description & vbCrLf & _
        "Source: " & Source
gKillLogging

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & gLogFileName & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' TradeBuild requires us to end the program without exiting from this event handler
End

End Sub

Public Sub gKillLogging()
GetLogger("").removeLogListener mListener
End Sub

Public Sub gModelessMsgBox( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
Dim lMsgBox As New fMsgBox

lMsgBox.initialise prompt, buttons, title

lMsgBox.Show vbModeless, gMainForm
                
End Sub

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

Public Sub gUnloadMainForm()
If Not mMainForm Is Nothing Then
    gLogger.Log LogLevelNormal, "Unloading main form"
    Unload mMainForm
    Set mMainForm = Nothing
End If
End Sub

Public Sub gUnsyncStudyPicker()
If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise Nothing, "Study picker"
End Sub

Public Sub Main()

Dim failpoint As Long

On Error GoTo Err


failpoint = 100

InitialiseTWUtilities
TaskConcurrency = 20
TaskQuantumMillisecs = 50

If showCommandLineOptions() Then Exit Sub


failpoint = 200

If Not getLog() Then Exit Sub

failpoint = 300

TradeBuildAPI.PermittedServiceProviderRoles = ServiceProviderRoles.SPRealtimeData Or _
                                                ServiceProviderRoles.SPPrimaryContractData Or _
                                                ServiceProviderRoles.SPSecondaryContractData Or _
                                                ServiceProviderRoles.SPBrokerLive Or _
                                                ServiceProviderRoles.SPBrokerSimulated Or _
                                                ServiceProviderRoles.SPHistoricalDataInput Or _
                                                ServiceProviderRoles.SPTickfileInput

If Not getConfigFile Then
    gLogger.Log LogLevelNormal, "Program exiting at user request"
    TerminateTWUtilities
ElseIf Not getConfig Then
    gLogger.Log LogLevelNormal, "Program exiting at user request"
    TerminateTWUtilities
ElseIf Not Configure Then
    gLogger.Log LogLevelNormal, "Program exiting at user request"
    TerminateTWUtilities
Else
    loadMainForm mEditConfig
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "main" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gHandleFatalError errNumber, errDescription, errSource
TerminateTWUtilities
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function Configure() As Boolean
Dim userResponse As Long

If ConfigureTradeBuild(gConfigFile, gAppInstanceConfig.InstanceQualifier) Then
    Configure = True
Else
    userResponse = MsgBox("The configuration cannot be loaded. Would you like to " & vbCrLf & _
            "manually correct the configuration?" & vbCrLf & vbCrLf & _
            "Click Yes to manually correct the configuration." & vbCrLf & vbCrLf & _
            "Click No to proceed with a new default configuration." & _
            "The default configuration will connect to TWS running on the " & vbCrLf & _
            "same computer. It will obtain contract data and historical data " & vbCrLf & _
            "from TWS, and will simulate any orders placed." & vbCrLf & vbCrLf & _
            "You may amend the default configuration by going to the " & vbCrLf & _
            "Configuration tab when the program has started." & vbCrLf & vbCrLf & _
            "Click Cancel to exit the program.", _
            vbYesNoCancel Or vbQuestion, _
            "Attention!")
    If userResponse = vbYes Then
        mEditConfig = True
        Configure = True
    ElseIf userResponse = vbNo Then
        gLogger.Log LogLevelNormal, "Creating a new default app instance configuration"
        Set gAppInstanceConfig = AddAppInstanceConfig(gConfigFile, _
                            DefaultAppInstanceConfigName, _
                            includeDefaultStudyLibrary:=True, _
                            setAsDefault:=True)
        Configure = True
    Else
        Configure = False
    End If
End If
End Function

Private Function getConfig() As Boolean
Dim configName As String

If gCommandLineParser.Switch(SwitchConfig) Then configName = gCommandLineParser.SwitchValue(SwitchConfig)

If configName = "" Then
    gLogger.Log LogLevelDetail, "Named config not specified - trying default config"
    configName = "(Default)"
    Set gAppInstanceConfig = GetDefaultAppInstanceConfig(gConfigFile)
    If gAppInstanceConfig Is Nothing Then
        gLogger.Log LogLevelDetail, "No default config defined"
    Else
        gLogger.Log LogLevelDetail, "Using default config: " & gAppInstanceConfig.InstanceQualifier
    End If
Else
    gLogger.Log LogLevelDetail, "Getting config with name '" & configName & "'"
    Set gAppInstanceConfig = GetAppInstanceConfig(gConfigFile, configName)
    If gAppInstanceConfig Is Nothing Then
        gLogger.Log LogLevelDetail, "Config '" & configName & "' not found"
    Else
        gLogger.Log LogLevelDetail, "Config '" & configName & "' located"
    End If
End If

If gAppInstanceConfig Is Nothing Then
    MsgBox "The required configuration does not exist: " & _
            configName & "." & vbCrLf & vbCrLf & _
            "The program will close.", _
            vbCritical, _
            "Error"
    getConfig = False
Else
    getConfig = True
End If

End Function

Private Function getConfigFile() As Boolean
Dim userResponse As Long
Dim baseConfigFile As TWUtilities30.configFile

On Error Resume Next
Set baseConfigFile = LoadXMLConfigurationFile(getConfigFilename)
On Error GoTo Err

If baseConfigFile Is Nothing Then
    userResponse = MsgBox("The configuration file does not exist." & vbCrLf & vbCrLf & _
            "Would you like to proceed with a default configuration?" & vbCrLf & vbCrLf & _
            "The default configuration will connect to TWS running on the " & vbCrLf & _
            "same computer. It will obtain contract data and historical data " & vbCrLf & _
            "from TWS." & vbCrLf & vbCrLf & _
            "You may amend the default configuration by going to the " & vbCrLf & _
            "Configuration tab when the program starts and using the " & vbCrLf & _
            "Configuration Editor." & vbCrLf & vbCrLf & _
            "Click Yes to continue with the default configuration." & vbCrLf & vbCrLf & _
            "Click No to exit the program", _
            vbYesNo Or vbQuestion, _
            "Attention!")
    If userResponse = vbYes Then
        gLogger.Log LogLevelNormal, "Creating a new default configuration file"
        Set baseConfigFile = CreateXMLConfigurationFile(App.ProductName, ConfigFileVersion)
        Set gConfigFile = CreateConfigurationFile(baseConfigFile, getConfigFilename)
        InitialiseConfigFile gConfigFile
        AddAppInstanceConfig gConfigFile, _
                            DefaultAppInstanceConfigName, _
                            includeDefaultStudyLibrary:=True, _
                            setAsDefault:=True
                            
    Else
        getConfigFile = False
        Exit Function
    End If
Else
    Set gConfigFile = CreateConfigurationFile(baseConfigFile, _
                                            getConfigFilename)
    If gConfigFile.applicationName <> App.ProductName Or _
        gConfigFile.fileVersion <> ConfigFileVersion Or _
        Not IsValidConfigurationFile(gConfigFile) _
    Then
        gLogger.Log LogLevelNormal, _
                    "The configuration file is not the correct format for this program." & vbCrLf & vbCrLf & _
                    "The program will close."
        getConfigFile = False
        Exit Function
    End If

End If

getConfigFile = True

Exit Function

Err:
gLogger.Log LogLevelNormal, "The configuration file format is not correct for this program."
MsgBox "The configuration file is not the correct format for this program" & vbCrLf & vbCrLf & _
        "The program will close."
End Function

Private Function getConfigFilename() As String

getConfigFilename = gCommandLineParser.Arg(0)
If getConfigFilename = "" Then
    getConfigFilename = gAppSettingsFolder & "\settings.xml"
End If
End Function

Private Function getLog() As Boolean

On Error GoTo Err

If gCommandLineParser.Switch(SwitchLogLevel) Then
    DefaultLogLevel = LogLevelFromString(gCommandLineParser.SwitchValue(SwitchLogLevel))
Else
    DefaultLogLevel = TWUtilities30.LogLevels.LogLevelNormal
End If

Set mListener = CreateFileLogListener(gLogFileName, _
                                        CreateBasicLogFormatter, _
                                        True, _
                                        False)
' ensure log entries of all infotypes get written to the log file
GetLogger("").addLogListener mListener

getLog = True
Exit Function

Err:
If Err.Number = ErrorCodes.ErrSecurityException Then
    MsgBox "You don't have write access to  '" & gLogFileName & "': the program will close", vbCritical, "Attention"
    getLog = False
Else
    Err.Raise Err.Number    ' unknown error so re-raise it
End If
End Function

Private Sub loadMainForm( _
                ByVal showConfigEditor As Boolean)
gLogger.Log LogLevelNormal, "Loading main form"
If mMainForm Is Nothing Then Set mMainForm = New fTradeSkilDemo
'mMainForm.Show
mMainForm.initialise showConfigEditor
mMainForm.Show
End Sub

Private Function showCommandLineOptions() As Boolean

If gCommandLineParser.Switch("?") Then
    MsgBox vbCrLf & _
            "tradeskildemo26 [configfile] " & vbCrLf & _
            "                [/config:configtoload] " & vbCrLf & _
            "                [/log:filename] " & vbCrLf & _
            "                [/loglevel:levelName]" & vbCrLf & _
            vbCrLf & _
            "  where" & vbCrLf & _
            vbCrLf & _
            "    levelname is one of:" & vbCrLf & _
            "       None    or 0" & vbCrLf & _
            "       Severe  or S" & vbCrLf & _
            "       Warning or W" & vbCrLf & _
            "       Info    or I" & vbCrLf & _
            "       Normal  or N" & vbCrLf & _
            "       Detail  or D" & vbCrLf & _
            "       Medium  or M" & vbCrLf & _
            "       High    or H" & vbCrLf & _
            "       All     or A", _
            , _
            "Usage"
    showCommandLineOptions = True
Else
    showCommandLineOptions = False
End If
End Function


