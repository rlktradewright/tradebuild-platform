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

Public Const WindowStateMaximized                   As String = "Maximized"
Public Const WindowStateMinimized                   As String = "Minimized"
Public Const WindowStateNormal                      As String = "Normal"

'@================================================================================
' Member variables
'@================================================================================

Private mIsInDev                                    As Boolean

Public gConfigFile                                  As ConfigurationFile
Public gAppInstanceConfig                           As ConfigurationSection

Private mStudyPickerForm                            As fStudyPicker
Private mMainForm                                   As fTradeSkilDemo

Private mEditConfig                                 As Boolean

Private mFatalErrorHandler                          As FatalErrorHandler

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
Const ProcName As String = "gCommandLineParser"
Static clp As CommandLineParser
Dim failpoint As String
On Error GoTo Err

If clp Is Nothing Then Set clp = CreateCommandLineParser(Command)
Set gCommandLineParser = clp

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get gMainForm() As fTradeSkilDemo
Set gMainForm = mMainForm
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gHandleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly. Note that normally one would use the
' End statement to terminate a VB6 program abruptly. However the TWUtilities component interferes
' with the End statement's processing and prevents proper shutdown, so we use the
' TWUtilities component's EndProcess method instead. (However if we are running in the
' development environment, then we call End because the EndProcess method kills the
' entire development environment as well which can have undesirable side effects if other
' components are also loaded.)

If mIsInDev Then
    ' this tells TWUtilities that we've now handled this unhandled error. Not actually
    ' needed here because the End statement will prevent return to TWUtilities
    UnhandledErrorHandler.Handled = True
    End
Else
    EndProcess
End If

End Sub

Public Sub gModelessMsgBox( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
Const ProcName As String = "gModelessMsgBox"
Dim lMsgBox As New fMsgBox

Dim failpoint As String
On Error GoTo Err

lMsgBox.initialise prompt, buttons, title

lMsgBox.Show vbModeless, gMainForm

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub gSetPermittedServiceProviderRoles()
Const ProcName As String = "gSetPermittedServiceProviderRoles"
Dim failpoint As String
On Error GoTo Err

TradeBuildAPI.PermittedServiceProviderRoles = ServiceProviderRoles.SPRealtimeData Or _
                                                ServiceProviderRoles.SPPrimaryContractData Or _
                                                ServiceProviderRoles.SPSecondaryContractData Or _
                                                ServiceProviderRoles.SPBrokerLive Or _
                                                ServiceProviderRoles.SPBrokerSimulated Or _
                                                ServiceProviderRoles.SPHistoricalDataInput Or _
                                                ServiceProviderRoles.SPTickfileInput

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub gShowStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal title As String)
Const ProcName As String = "gShowStudyPicker"
Dim failpoint As String
On Error GoTo Err

If mStudyPickerForm Is Nothing Then Set mStudyPickerForm = New fStudyPicker
mStudyPickerForm.initialise chartMgr, title
mStudyPickerForm.Show vbModeless

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub gSyncStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal title As String)
Const ProcName As String = "gSyncStudyPicker"
Dim failpoint As String
On Error GoTo Err

If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise chartMgr, title

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub gUnloadMainForm()
Const ProcName As String = "gUnloadMainForm"
Dim failpoint As String
On Error GoTo Err

If Not mMainForm Is Nothing Then
    LogMessage "Unloading main form"
    Unload mMainForm
    Set mMainForm = Nothing
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub gUnsyncStudyPicker()
Const ProcName As String = "gUnsyncStudyPicker"
Dim failpoint As String
On Error GoTo Err

If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise Nothing, "Study picker"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub Main()
Const ProcName As String = "Main"
Dim failpoint As String
On Error GoTo Err

failpoint = 100

Debug.Print "Running in development environment: " & CStr(inDev)

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler

If showCommandLineOptions() Then Exit Sub

ApplicationGroupName = "TradeWright"
applicationName = gAppTitle
SetupDefaultLogging Command

TaskConcurrency = 20
TaskQuantumMillisecs = 32

failpoint = 300

gSetPermittedServiceProviderRoles

If Not getConfigFile Then
    LogMessage "Program exiting at user request"
    TerminateTWUtilities
ElseIf Not getConfig Then
    LogMessage "Program exiting at user request"
    TerminateTWUtilities
ElseIf Not Configure Then
    LogMessage "Program exiting at user request"
    TerminateTWUtilities
Else
    loadMainForm mEditConfig
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrSecurityException Then
    MsgBox "You don't have write access to the log file:" & vbCrLf & vbCrLf & _
                DefaultLogFileName & vbCrLf & vbCrLf & _
                "The program will close", _
            vbCritical, _
            "Attention"
    Exit Sub
End If
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function Configure() As Boolean
Const ProcName As String = "Configure"
Dim userResponse As Long

Dim failpoint As String
On Error GoTo Err

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
        LogMessage "Creating a new default app instance configuration"
        Set gAppInstanceConfig = AddAppInstanceConfig(gConfigFile, _
                            DefaultAppInstanceConfigName, _
                            includeDefaultStudyLibrary:=True, _
                            setAsDefault:=True)
        Configure = True
    Else
        Configure = False
    End If
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function getConfig() As Boolean
Const ProcName As String = "getConfig"
Dim configName As String

Dim failpoint As String
On Error GoTo Err

If gCommandLineParser.Switch(SwitchConfig) Then configName = gCommandLineParser.SwitchValue(SwitchConfig)

If configName = "" Then
    LogMessage "Named config not specified - trying default config", LogLevelDetail
    configName = "(Default)"
    Set gAppInstanceConfig = GetDefaultAppInstanceConfig(gConfigFile)
    If gAppInstanceConfig Is Nothing Then
        LogMessage "No default config defined", LogLevelDetail
    Else
        LogMessage "Using default config: " & gAppInstanceConfig.InstanceQualifier, LogLevelDetail
    End If
Else
    LogMessage "Getting config with name '" & configName & "'", LogLevelDetail
    Set gAppInstanceConfig = GetAppInstanceConfig(gConfigFile, configName)
    If gAppInstanceConfig Is Nothing Then
        LogMessage "Config '" & configName & "' not found"
    Else
        LogMessage "Config '" & configName & "' located", LogLevelDetail
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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Function getConfigFile() As Boolean
Const ProcName As String = "getConfigFile"
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
        LogMessage "Creating a new default configuration file"
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
        LogMessage "The configuration file is not the correct format for this program." & vbCrLf & vbCrLf & _
                    "The program will close."
        getConfigFile = False
        Exit Function
    End If

End If

getConfigFile = True

Exit Function

Err:
LogMessage "The configuration file format is not correct for this program."
MsgBox "The configuration file is not the correct format for this program" & vbCrLf & vbCrLf & _
        "The program will close."
End Function

Private Function getConfigFilename() As String
Const ProcName As String = "getConfigFilename"


Dim failpoint As String
On Error GoTo Err

getConfigFilename = gCommandLineParser.Arg(0)
If getConfigFilename = "" Then
    getConfigFilename = ApplicationSettingsFolder & "\settings.xml"
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function inDev() As Boolean
Const ProcName As String = "inDev"
On Error GoTo Err

mIsInDev = True
inDev = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub loadMainForm( _
                ByVal showConfigEditor As Boolean)
Const ProcName As String = "loadMainForm"
Dim failpoint As String
On Error GoTo Err

LogMessage "Loading main form"
If mMainForm Is Nothing Then Set mMainForm = New fTradeSkilDemo
'mMainForm.Show
mMainForm.initialise showConfigEditor
mMainForm.Show

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function showCommandLineOptions() As Boolean
Const ProcName As String = "showCommandLineOptions"


Dim failpoint As String
On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function


