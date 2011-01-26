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

Public Const ConfigFileVersion                      As String = "1.2"

Public Const ConfigSectionApplication               As String = "Application"
Public Const ConfigSectionChart                     As String = "Chart"
Public Const ConfigSectionCharts                    As String = "Charts"
Public Const ConfigSectionConfigEditor              As String = "ConfigEditor"
Public Const ConfigSectionDefaultStudyConfigs       As String = "DefaultStudyConfigs"
Public Const ConfigSectionChartStyles               As String = "/ChartStyles"
Public Const ConfigSectionMainForm                  As String = "MainForm"
Public Const ConfigSectionMultiChart                As String = "MultiChart"
Public Const ConfigSectionOrderTicket               As String = "OrderTicket"
Public Const ConfigSectionTickerGrid                As String = "TickerGrid"

Public Const ConfigSettingHeight                    As String = ".Height"
Public Const ConfigSettingLeft                      As String = ".Left"
Public Const ConfigSettingTop                       As String = ".Top"
Public Const ConfigSettingWidth                     As String = ".Width"
Public Const ConfigSettingWindowstate               As String = ".Windowstate"

Public Const ConfigSettingCurrentChartStyle         As String = "&CurrentChartStyle"

Public Const ConfigSettingAppCurrentChartStyle      As String = ConfigSectionApplication & ConfigSettingCurrentChartStyle

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

Public Const InitialChartStyleName                  As String = "Application default"

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

Public gConfigStore                                  As ConfigurationStore
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

On Error GoTo Err

If clp Is Nothing Then Set clp = CreateCommandLineParser(Command)
Set gCommandLineParser = clp

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get gMainForm() As fTradeSkilDemo
Set gMainForm = mMainForm
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gHandleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

' kill off any timers
TerminateTWUtilities

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly.
'
' Note that normally one would use the End statement to terminate a VB6 program abruptly. But
' the TWUtilities component interferes with the End statement's processing and may prevent
' proper shutdown, so we use the TWUtilities component's EndProcess method instead.
'
' However if we are running in the development environment, then we call End because the
' EndProcess method kills the entire development environment as well which can have undesirable
' side effects if other components are also loaded.

If mIsInDev Then
    End
Else
    EndProcess
End If

End Sub

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub gModelessMsgBox( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
Const ProcName As String = "gModelessMsgBox"
Dim lMsgBox As New fMsgBox


On Error GoTo Err

lMsgBox.initialise prompt, buttons, title

lMsgBox.Show vbModeless, gMainForm

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub gSetPermittedServiceProviderRoles()
Const ProcName As String = "gSetPermittedServiceProviderRoles"

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub gShowStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal title As String)
Const ProcName As String = "gShowStudyPicker"

On Error GoTo Err

If mStudyPickerForm Is Nothing Then Set mStudyPickerForm = New fStudyPicker
mStudyPickerForm.initialise chartMgr, title
mStudyPickerForm.Show vbModeless

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub gSyncStudyPicker( _
                ByVal chartMgr As ChartManager, _
                ByVal title As String)
Const ProcName As String = "gSyncStudyPicker"

On Error GoTo Err

If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise chartMgr, title

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub gUnloadMainForm()
Const ProcName As String = "gUnloadMainForm"

On Error GoTo Err

If Not mMainForm Is Nothing Then
    LogMessage "Unloading main form"
    Unload mMainForm
    Set mMainForm = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub gUnsyncStudyPicker()
Const ProcName As String = "gUnsyncStudyPicker"

On Error GoTo Err

If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise Nothing, "Study picker"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Main()
Const ProcName As String = "Main"

Dim lSplash As fSplash


On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler

If showCommandLineOptions() Then Exit Sub

ApplicationGroupName = "TradeWright"
ApplicationName = gAppTitle
SetupDefaultLogging Command

TaskConcurrency = 20
TaskQuantumMillisecs = 32

Set lSplash = showSplashScreen

gSetPermittedServiceProviderRoles

If Not getConfigFile Then
    LogMessage "Program exiting at user request"
    TerminateTWUtilities
ElseIf Not getAppInstanceConfig Then
    LogMessage "Program exiting at user request"
    TerminateTWUtilities
ElseIf Not Configure Then
    LogMessage "Program exiting at user request"
    TerminateTWUtilities
Else
    loadMainForm mEditConfig
End If

Unload lSplash
Exit Sub

Err:
If Err.Number = ErrorCodes.ErrSecurityException Then
    MsgBox "You don't have write access to the log file:" & vbCrLf & vbCrLf & _
                DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
                "The program will close", _
            vbCritical, _
            "Attention"
    Exit Sub
End If
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function Configure() As Boolean
Const ProcName As String = "Configure"
Dim userResponse As Long


On Error GoTo Err

LogMessage "Loading configuration: loading chart styles"
ChartStylesManager.LoadFromConfig gConfigStore.AddPrivateConfigurationSection(ConfigSectionChartStyles)
    
If ConfigureTradeBuild(gConfigStore, gAppInstanceConfig.InstanceQualifier) Then
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
        Set gAppInstanceConfig = AddAppInstanceConfig(gConfigStore, _
                            DefaultAppInstanceConfigName, _
                            ConfigFlagIncludeDefaultBarFormatterLibrary Or _
                                ConfigFlagIncludeDefaultStudyLibrary Or _
                                ConfigFlagSetAsDefault)
        Configure = True
    Else
        Configure = False
    End If
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function createNewConfigFile() As ConfigurationStore
Const ProcName As String = "createNewConfigFile"
On Error GoTo Err

LogMessage "Creating a new default configuration file"
Set createNewConfigFile = GetDefaultConfigurationStore(Command, ConfigFileVersion, True, ConfigFileOptionFirstArg)
InitialiseConfigFile createNewConfigFile
AddAppInstanceConfig createNewConfigFile, _
                    DefaultAppInstanceConfigName, _
                    ConfigFlagIncludeDefaultBarFormatterLibrary Or _
                        ConfigFlagIncludeDefaultStudyLibrary Or _
                        ConfigFlagSetAsDefault

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getAppInstanceConfig() As Boolean
Const ProcName As String = "getAppInstanceConfig"
Dim configName As String


On Error GoTo Err

If gCommandLineParser.Switch(SwitchConfig) Then configName = gCommandLineParser.SwitchValue(SwitchConfig)

If configName = "" Then
    LogMessage "Named app instance config not specified - trying default app instance config", LogLevelDetail
    configName = "(Default)"
    Set gAppInstanceConfig = GetDefaultAppInstanceConfig(gConfigStore)
    If gAppInstanceConfig Is Nothing Then
        LogMessage "No default app instance config defined", LogLevelDetail
    Else
        LogMessage "Using default app instance config: " & gAppInstanceConfig.InstanceQualifier, LogLevelDetail
    End If
Else
    LogMessage "Getting app instance config with name '" & configName & "'", LogLevelDetail
    Set gAppInstanceConfig = ConfigUtils.getAppInstanceConfig(gConfigStore, configName)
    If gAppInstanceConfig Is Nothing Then
        LogMessage "App instance config '" & configName & "' not found"
    Else
        LogMessage "App instance config '" & configName & "' located", LogLevelDetail
    End If
End If

If gAppInstanceConfig Is Nothing Then
    MsgBox "The required app instance configuration does not exist: " & _
            configName & "." & vbCrLf & vbCrLf & _
            "The program will close.", _
            vbCritical, _
            "Error"
    getAppInstanceConfig = False
Else
    getAppInstanceConfig = True
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getConfigFile() As Boolean
Const ProcName As String = "getConfigFile"

On Error Resume Next
Set gConfigStore = GetDefaultConfigurationStore(Command, ConfigFileVersion, False, ConfigFileOptionFirstArg)

If Err.Number = ErrorCodes.ErrIllegalStateException Then
    On Error GoTo Err
    
    getConfigFile = queryReplaceConfigFile
ElseIf gConfigStore Is Nothing Then
    On Error GoTo Err
    
    getConfigFile = queryCreateNewConfigFile
ElseIf IsValidConfigurationFile(gConfigStore) Then
    On Error GoTo Err
    getConfigFile = True
Else
    On Error GoTo Err
    
    getConfigFile = queryReplaceConfigFile
End If



Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function inDev() As Boolean
Const ProcName As String = "inDev"
On Error GoTo Err

mIsInDev = True
inDev = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub loadMainForm( _
                ByVal showConfigEditor As Boolean)
Const ProcName As String = "loadMainForm"

On Error GoTo Err

LogMessage "Loading main form"
If mMainForm Is Nothing Then Set mMainForm = New fTradeSkilDemo
mMainForm.initialise showConfigEditor
mMainForm.Show

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function queryCreateNewConfigFile() As Boolean
Const ProcName As String = "queryCreateNewConfigFile"
On Error GoTo Err

Dim userResponse As Long
LogMessage "The configuration file format does not exist."
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
    Set gConfigStore = createNewConfigFile
    queryCreateNewConfigFile = True
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function queryReplaceConfigFile() As Boolean
Const ProcName As String = "queryReplaceConfigFile"
On Error GoTo Err

Dim userResponse As Long
LogMessage "The configuration file format is not correct for this program."
userResponse = MsgBox("The configuration file is not the correct format for this program." & vbCrLf & vbCrLf & _
        "This may be because you have installed a new version of " & vbCrLf & _
        "the program, or because the file has been corrupted." & vbCrLf & vbCrLf & _
        "Would you like to proceed with a default configuration?" & vbCrLf & vbCrLf & _
        "The default configuration will connect to TWS running on the " & vbCrLf & _
        "same computer. It will obtain contract data and historical data " & vbCrLf & _
        "from TWS." & vbCrLf & vbCrLf & _
        "You may amend the default configuration by going to the " & vbCrLf & _
        "Configuration tab when the program starts and using the " & vbCrLf & _
        "Configuration Editor." & vbCrLf & vbCrLf & _
        "Note that the default configuration will overwrite your " & vbCrLf & _
        "current configuration file and all settings in it will be " & vbCrLf & _
        "lost." & vbCrLf & vbCrLf & _
        "Click Yes (recommended) to continue with the default configuration." & vbCrLf & vbCrLf & _
        "Click No to exit the program", _
        vbYesNo Or vbQuestion, _
        "Attention!")
If userResponse = vbYes Then
    Set gConfigStore = createNewConfigFile
    queryReplaceConfigFile = True
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function showCommandLineOptions() As Boolean
Const ProcName As String = "showCommandLineOptions"



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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function showSplashScreen() As Form
Dim lSplash As New fSplash
Const ProcName As String = "showSplashScreen"
On Error GoTo Err

lSplash.Show vbModeless
lSplash.Refresh
Set showSplashScreen = lSplash
SetWindowLong lSplash.hWnd, GWL_EXSTYLE, GetWindowLong(lSplash.hWnd, GWL_EXSTYLE) Or WS_EX_TOPMOST
SetWindowPos lSplash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function
