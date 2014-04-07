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

Public Const ProjectName                            As String = "TradeSkilDemo27"
Private Const ModuleName                            As String = "GMain"

Public Const AppName                                As String = "TradeSkil Demo Edition"

Public Const ConfigFileVersion                      As String = "1.3"

Public Const ConfigSectionApplication               As String = "Application"
Public Const ConfigSectionChart                     As String = "Chart"
Public Const ConfigSectionCharts                    As String = "Charts"
Public Const ConfigSectionConfigEditor              As String = "ConfigEditor"
Public Const ConfigSectionContract                  As String = "Contract"
Public Const ConfigSectionDefaultStudyConfigs       As String = "DefaultStudyConfigs"
Public Const ConfigSectionChartStyles               As String = "/ChartStyles"
Public Const ConfigSectionHistoricCharts            As String = "HistoricCharts"
Public Const ConfigSectionMainForm                  As String = "MainForm"
Public Const ConfigSectionMultiChart                As String = "MultiChart"
Public Const ConfigSectionOrderTicket               As String = "OrderTicket"
Public Const ConfigSectionTickerGrid                As String = "TickerGrid"

Public Const ConfigSettingHistorical                As String = "&Historical"
Public Const ConfigSettingHeight                    As String = "&Height"
Public Const ConfigSettingLeft                      As String = "&Left"
Public Const ConfigSettingTop                       As String = "&Top"
Public Const ConfigSettingWidth                     As String = "&Width"
Public Const ConfigSettingWindowstate               As String = "&Windowstate"

Public Const ConfigSettingCurrentChartStyle         As String = "&CurrentChartStyle"
Public Const ConfigSettingCurrentHistChartStyle     As String = "&CurrentHistChartStyle"

Public Const ConfigSettingAppCurrentChartStyle      As String = ConfigSectionApplication & ConfigSettingCurrentChartStyle
Public Const ConfigSettingAppCurrentHistChartStyle  As String = ConfigSectionApplication & ConfigSettingCurrentHistChartStyle

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

Public Const ChartStyleNameAppDefault               As String = "Application default"
Public Const ChartStyleNameBlack                    As String = "Black"
Public Const ChartStyleNameDarkBlueFade             As String = "Dark blue fade"
Public Const ChartStyleNameGoldFade                 As String = "Gold fade"

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

Public mConfigStore                                 As ConfigurationStore
'Public gAppInstanceConfig                           As ConfigurationSection

Private mEditConfig                                 As Boolean

Private mFatalErrorHandler                          As FatalErrorHandler

Private mMainForm                                   As fTradeSkilDemo

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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get gMainForm() As fTradeSkilDemo
Set gMainForm = mMainForm
End Property

'@================================================================================
' Methods
'@================================================================================

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

Public Function gLoadMainForm( _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                Optional ByVal pPrevMainForm As fTradeSkilDemo, _
                Optional ByVal pSplash As fSplash) As Boolean
Const ProcName As String = "loadMainForm"
On Error GoTo Err

Dim lMainForm As New fTradeSkilDemo

LogMessage "Configuring TradeBuild with config: " & pAppInstanceConfig.InstanceQualifier
Dim lTradeBuildAPI As TradeBuildAPI
Set lTradeBuildAPI = configureTradeBuild(mConfigStore, pAppInstanceConfig)

If lTradeBuildAPI Is Nothing Then
    LogMessage "Failed to configure TradeBuild with config: " & pAppInstanceConfig.InstanceQualifier
    gLoadMainForm = False
    Exit Function
End If

LogMessage "Successfully configured TradeBuild with config: " & pAppInstanceConfig.InstanceQualifier
LogMessage "Loading main form for config: " & pAppInstanceConfig.InstanceQualifier

lMainForm.Initialise lTradeBuildAPI, mConfigStore, pAppInstanceConfig, pSplash, pPrevMainForm

LogMessage "Main form initialised successfully"

Set mMainForm = lMainForm
gLoadMainForm = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gModelessMsgBox( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
Const ProcName As String = "gModelessMsgBox"
On Error GoTo Err

Dim lMsgBox As New fMsgBox
lMsgBox.Initialise prompt, buttons, title
lMsgBox.Show vbModeless, gMainForm

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Property Get gPermittedServiceProviderRoles() As ServiceProviderRoles
gPermittedServiceProviderRoles = ServiceProviderRoles.SPRoleRealtimeData Or _
                                ServiceProviderRoles.SPRoleContractDataPrimary Or _
                                ServiceProviderRoles.SPRoleContractDataSecondary Or _
                                ServiceProviderRoles.SPRoleOrderPersistence Or _
                                ServiceProviderRoles.SPRoleOrderSubmissionLive Or _
                                ServiceProviderRoles.SPRoleOrderSubmissionSimulated Or _
                                ServiceProviderRoles.SPRoleHistoricalDataInput Or _
                                ServiceProviderRoles.SPRoleTickfileInput
End Property

Public Sub gSaveSettings()
Const ProcName As String = "gSaveSettings"
On Error GoTo Err

If mConfigStore.Dirty Then
    LogMessage "Saving configuration"
    mConfigStore.Save
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gShowConfigEditor( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pCurrAppInstanceConfig As ConfigurationSection, _
                Optional ByVal pParentForm As Form) As ConfigurationSection
Const ProcName As String = "gShowConfigEditor"
On Error GoTo Err

Dim lConfigEditor As New fConfigEditor
   
lConfigEditor.Initialise pConfigStore, pCurrAppInstanceConfig
lConfigEditor.Show vbModal, pParentForm

Set gShowConfigEditor = lConfigEditor.selectedAppConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gShowSplashScreen() As Form
Const ProcName As String = "gShowSplashScreen"
On Error GoTo Err

Dim lSplash As New fSplash
lSplash.Show vbModeless
lSplash.Refresh
Set gShowSplashScreen = lSplash
SetWindowLong lSplash.hWnd, GWL_EXSTYLE, GetWindowLong(lSplash.hWnd, GWL_EXSTYLE) Or WS_EX_TOPMOST
SetWindowPos lSplash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'Public Sub gShowStudyPicker( _
'                ByVal chartMgr As ChartManager, _
'                ByVal title As String)
'Const ProcName As String = "gShowStudyPicker"
'
'On Error GoTo Err
'
'If mStudyPickerForm Is Nothing Then Set mStudyPickerForm = New fStudyPicker
'mStudyPickerForm.Initialise chartMgr, title
'mStudyPickerForm.Show vbModeless, mMainForm
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'Public Sub gSyncStudyPicker( _
'                ByVal chartMgr As ChartManager, _
'                ByVal title As String)
'Const ProcName As String = "gSyncStudyPicker"
'
'On Error GoTo Err
'
'If mStudyPickerForm Is Nothing Then Exit Sub
'mStudyPickerForm.Initialise chartMgr, title
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

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
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Public Sub gUnsyncStudyPicker()
'Const ProcName As String = "gUnsyncStudyPicker"
'
'On Error GoTo Err
'
'If mStudyPickerForm Is Nothing Then Exit Sub
'mStudyPickerForm.Initialise Nothing, "Study picker"
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Public Sub Main()
Const ProcName As String = "Main"
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

Dim lSplash As fSplash
Set lSplash = gShowSplashScreen

Set mConfigStore = getConfigStore
If mConfigStore Is Nothing Then
    LogMessage "Program exiting at user request"
    Unload lSplash
    TerminateTWUtilities
    Exit Sub
End If

loadChartStyles mConfigStore
    
Dim lAppInstanceConfig As ConfigurationSection
Set lAppInstanceConfig = getAppInstanceConfig(mConfigStore)
If lAppInstanceConfig Is Nothing Then
    LogMessage "Program exiting at user request"
    Unload lSplash
    TerminateTWUtilities
    Exit Sub
End If

If Not gLoadMainForm(lAppInstanceConfig, , lSplash) Then
    LogMessage "Program exiting at user request"
    Unload lSplash
    TerminateTWUtilities
    Exit Sub
End If

Do While Forms.Count > 0
    Wait 50
Loop

gSaveSettings

TerminateTWUtilities

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrSecurityException Then
    MsgBox "You don't have write access to the log file:" & vbCrLf & vbCrLf & _
                DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
                "The program will close", _
            vbCritical, _
            "Attention"
    TerminateTWUtilities
    Exit Sub
End If
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function configureTradeBuild( _
                ByVal pConfigStore As ConfigurationStore, _
                ByRef pAppInstanceConfig As ConfigurationSection) As TradeBuildAPI
Const ProcName As String = "configureTradeBuild"
On Error GoTo Err

Dim lTradeBuildAPI As TradeBuildAPI
Set lTradeBuildAPI = CreateTradeBuildAPIFromConfig(pAppInstanceConfig, pAppInstanceConfig.InstanceQualifier, gPermittedServiceProviderRoles)
Do While lTradeBuildAPI Is Nothing
    Dim userResponse As Long
    userResponse = MsgBox("The configuration cannot be loaded. Would you like to " & vbCrLf & _
            "manually correct the configuration?" & vbCrLf & vbCrLf & _
            "Click Yes to manually correct the configuration." & vbCrLf & vbCrLf & _
            "Click No to proceed with a new default configuration." & _
            "The default configuration will connect to TWS running on the " & vbCrLf & _
            "same computer. It will obtain contract data and historical data " & vbCrLf & _
            "from TWS, and will simulate any orders placed." & vbCrLf & vbCrLf & _
            "You may amend the default configuration by going to the " & vbCrLf & _
            "Configuration tab when the program has started." & vbCrLf & vbCrLf & _
            "Click Cancel to exit.", _
            vbYesNoCancel Or vbQuestion, _
            "Attention!")
    If userResponse = vbYes Then
        LogMessage "User editing app instance configuration: " & pAppInstanceConfig.InstanceQualifier
        Set pAppInstanceConfig = gShowConfigEditor(pConfigStore, pAppInstanceConfig)
    ElseIf userResponse = vbNo Then
        LogMessage "Creating a new default app instance configuration"
        Set pAppInstanceConfig = AddAppInstanceConfig(pConfigStore, _
                            DefaultAppInstanceConfigName, _
                            ConfigFlagIncludeDefaultBarFormatterLibrary Or _
                                ConfigFlagIncludeDefaultStudyLibrary Or _
                                ConfigFlagSetAsDefault)
    Else
        Exit Do
    End If
        
    Set lTradeBuildAPI = CreateTradeBuildAPIFromConfig(pAppInstanceConfig, pAppInstanceConfig.InstanceQualifier, gPermittedServiceProviderRoles)
Loop

Assert lTradeBuildAPI.StartServiceProviders, "Error while starting service providers"
Set configureTradeBuild = lTradeBuildAPI

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createNewConfigStore() As ConfigurationStore
Const ProcName As String = "createNewConfigStore"
On Error GoTo Err

LogMessage "Creating a new default configuration file"

Dim lConfigStore As ConfigurationStore
Set lConfigStore = GetDefaultConfigurationStore(Command, ConfigFileVersion, True, ConfigFileOptionFirstArg)
InitialiseConfigFile lConfigStore
AddAppInstanceConfig lConfigStore, _
                    DefaultAppInstanceConfigName, _
                    ConfigFlagIncludeDefaultBarFormatterLibrary Or _
                        ConfigFlagIncludeDefaultStudyLibrary Or _
                        ConfigFlagSetAsDefault, _
                    gPermittedServiceProviderRoles
lConfigStore.Save

Set createNewConfigStore = lConfigStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getAppInstanceConfig(ByVal pConfigStore As ConfigurationStore) As ConfigurationSection
Const ProcName As String = "getAppInstanceConfig"
On Error GoTo Err

Dim lAppInstanceConfig As ConfigurationSection

Dim configName As String
If gCommandLineParser.Switch(SwitchConfig) Then configName = gCommandLineParser.SwitchValue(SwitchConfig)

If configName = "" Then
    LogMessage "Named app instance config not specified - trying default app instance config", LogLevelDetail
    configName = "(Default)"
    Set lAppInstanceConfig = GetDefaultAppInstanceConfig(pConfigStore)
    If lAppInstanceConfig Is Nothing Then
        LogMessage "No default app instance config defined", LogLevelDetail
    Else
        LogMessage "Using default app instance config: " & lAppInstanceConfig.InstanceQualifier, LogLevelDetail
    End If
Else
    LogMessage "Getting app instance config with name '" & configName & "'", LogLevelDetail
    Set lAppInstanceConfig = ConfigUtils.getAppInstanceConfig(pConfigStore, configName)
    If lAppInstanceConfig Is Nothing Then
        LogMessage "App instance config '" & configName & "' not found"
    Else
        LogMessage "App instance config '" & configName & "' located", LogLevelDetail
    End If
End If

If lAppInstanceConfig Is Nothing Then
    MsgBox "The required app instance configuration does not exist: " & _
            configName & "." & vbCrLf & vbCrLf & _
            "The program will close.", _
            vbCritical, _
            "Error"
End If

Set getAppInstanceConfig = lAppInstanceConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getConfigStore() As ConfigurationStore
Const ProcName As String = "getConfigStore"
On Error GoTo Err

On Error Resume Next

Dim lConfigStore As ConfigurationStore
Set lConfigStore = GetDefaultConfigurationStore(Command, ConfigFileVersion, False, ConfigFileOptionFirstArg)

If Err.Number = ErrorCodes.ErrIllegalStateException Then
    On Error GoTo Err
    If queryReplaceConfigFile Then Set lConfigStore = createNewConfigStore
ElseIf lConfigStore Is Nothing Then
    On Error GoTo Err
    If queryCreateNewConfigFile Then Set lConfigStore = createNewConfigStore
ElseIf Not IsValidConfigurationFile(lConfigStore) Then
    On Error GoTo Err
    If queryReplaceConfigFile Then Set lConfigStore = createNewConfigStore
End If

Set getConfigStore = lConfigStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function inDev() As Boolean
Const ProcName As String = "inDev"
On Error GoTo Err

mIsInDev = True
inDev = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadChartStyles(ByVal pConfigStore As ConfigurationStore)
LogMessage "Loading configuration: loading chart styles"
ChartStylesManager.LoadFromConfig pConfigStore.AddPrivateConfigurationSection(ConfigSectionChartStyles)
End Sub

Private Function queryCreateNewConfigFile() As Boolean
Const ProcName As String = "queryCreateNewConfigFile"
On Error GoTo Err

Dim userResponse As Long
LogMessage "The configuration file does not exist."
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
    queryCreateNewConfigFile = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    queryReplaceConfigFile = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function showCommandLineOptions() As Boolean
Const ProcName As String = "showCommandLineOptions"
On Error GoTo Err

If gCommandLineParser.Switch("?") Then
    MsgBox vbCrLf & _
            "tradeskildemo27 [configfile] " & vbCrLf & _
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
gHandleUnexpectedError ProcName, ModuleName
End Function

