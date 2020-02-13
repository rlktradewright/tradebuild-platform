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
'Public Const ConfigSectionChart                     As String = "Chart"
Public Const ConfigSectionCharts                    As String = "Charts"
Public Const ConfigSectionConfigEditor              As String = "ConfigEditor"
'Public Const ConfigSectionContract                  As String = "Contract"
Public Const ConfigSectionDefaultStudyConfigs       As String = "DefaultStudyConfigs"
Public Const ConfigSectionChartStyles               As String = "/ChartStyles"
Public Const ConfigSectionFloatingFeaturesPanel     As String = "FloatingFeaturesPanel"
Public Const ConfigSectionFloatingInfoPanel         As String = "FloatingInfoPanel"
Public Const ConfigSectionHistoricCharts            As String = "HistoricCharts"
Public Const ConfigSectionMainForm                  As String = "MainForm"
Public Const ConfigSectionOrderTicket               As String = "OrderTicket"
Public Const ConfigSectionTickerGrid                As String = "TickerGrid"

Public Const ConfigSettingDataSourceKey             As String = "&DataSourceKey"
Public Const ConfigSettingHistorical                As String = "&Historical"
Public Const ConfigSettingHeight                    As String = "&Height"
Public Const ConfigSettingLeft                      As String = "&Left"
Public Const ConfigSettingTop                       As String = "&Top"
Public Const ConfigSettingWidth                     As String = "&Width"
Public Const ConfigSettingWindowstate               As String = "&Windowstate"
Public Const ConfigSettingFeaturesPanelHidden       As String = "&FeaturesPanelHidden"
Public Const ConfigSettingFeaturesPanelPinned       As String = "&FeaturesPanelPinned"
Public Const ConfigSettingInfoPanelHidden           As String = "&InfoPanelHidden"
Public Const ConfigSettingInfoPanelPinned           As String = "&InfoPanelPinned"
Public Const ConfigSettingCurrentTheme              As String = "&CurrentTheme"

Public Const ConfigSettingCurrentChartStyle         As String = "&CurrentChartStyle"
Public Const ConfigSettingCurrentHistChartStyle     As String = "&CurrentHistChartStyle"

Public Const ConfigSettingAppCurrentChartStyle      As String = ConfigSectionApplication & ConfigSettingCurrentChartStyle
Public Const ConfigSettingAppCurrentHistChartStyle  As String = ConfigSectionApplication & ConfigSettingCurrentHistChartStyle

Public Const ConfigSettingConfigEditorLeft          As String = ConfigSectionConfigEditor & ConfigSettingLeft
Public Const ConfigSettingConfigEditorTop           As String = ConfigSectionConfigEditor & ConfigSettingTop

Public Const ConfigSettingMainFormHeight            As String = ConfigSectionMainForm & ConfigSettingHeight
Public Const ConfigSettingMainFormLeft              As String = ConfigSectionMainForm & ConfigSettingLeft
Public Const ConfigSettingMainFormTop               As String = ConfigSectionMainForm & ConfigSettingTop
Public Const ConfigSettingMainFormWidth             As String = ConfigSectionMainForm & ConfigSettingWidth
Public Const ConfigSettingMainFormWindowstate       As String = ConfigSectionMainForm & ConfigSettingWindowstate

Public Const ConfigSettingOrderTicketLeft           As String = ConfigSectionOrderTicket & ConfigSettingLeft
Public Const ConfigSettingOrderTicketTop            As String = ConfigSectionOrderTicket & ConfigSettingTop

Public Const ConfigSettingFloatingFeaturesPanelLeft As String = ConfigSectionFloatingFeaturesPanel & ConfigSettingLeft
Public Const ConfigSettingFloatingFeaturesPanelTop  As String = ConfigSectionFloatingFeaturesPanel & ConfigSettingTop
Public Const ConfigSettingFloatingFeaturesPanelWidth  As String = ConfigSectionFloatingFeaturesPanel & ConfigSettingWidth
Public Const ConfigSettingFloatingFeaturesPanelHeight As String = ConfigSectionFloatingFeaturesPanel & ConfigSettingHeight

Public Const ConfigSettingFloatingInfoPanelLeft     As String = ConfigSectionFloatingInfoPanel & ConfigSettingLeft
Public Const ConfigSettingFloatingInfoPanelTop      As String = ConfigSectionFloatingInfoPanel & ConfigSettingTop
Public Const ConfigSettingFloatingInfoPanelWidth    As String = ConfigSectionFloatingInfoPanel & ConfigSettingWidth
Public Const ConfigSettingFloatingInfoPanelHeight   As String = ConfigSectionFloatingInfoPanel & ConfigSettingHeight

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

Private mConfigStore                                As ConfigurationStore
Private mConfigChangeMonitor                        As ConfigChangeMonitor

Private mEditConfig                                 As Boolean

Private mFatalErrorHandler                          As FatalErrorHandler

Private mMainForm                                   As fTradeSkilDemo

Private mConfigEditor                               As fConfigEditor

Private mSplash                                     As fSplash

Private mFinished                                   As Boolean

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

Public Property Get gSplashScreen() As fSplash
Set gSplashScreen = mSplash
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gApplyTheme(ByVal pTheme As ITheme, ByVal pControls As Object)
Const ProcName As String = "gApplyTheme"
On Error GoTo Err

If pTheme Is Nothing Then Exit Sub

Dim lControl As Control
For Each lControl In pControls
    If TypeOf lControl Is Label Then
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is CheckBox Or _
        TypeOf lControl Is Frame Or _
        TypeOf lControl Is OptionButton _
    Then
        SetWindowThemeOff lControl.hWnd
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is PictureBox Then
        lControl.Appearance = pTheme.Appearance
        lControl.BorderStyle = pTheme.BorderStyle
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is TextBox Then
        lControl.Appearance = pTheme.Appearance
        lControl.BorderStyle = pTheme.BorderStyle
        lControl.BackColor = pTheme.TextBackColor
        lControl.ForeColor = pTheme.TextForeColor
        If Not pTheme.TextFont Is Nothing Then
            Set lControl.Font = pTheme.TextFont
        ElseIf Not pTheme.BaseFont Is Nothing Then
            Set lControl.Font = pTheme.BaseFont
        End If
    ElseIf TypeOf lControl Is ComboBox Or _
        TypeOf lControl Is ListBox _
    Then
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.TextBackColor
        lControl.ForeColor = pTheme.TextForeColor
        If Not pTheme.ComboFont Is Nothing Then
            Set lControl.Font = pTheme.ComboFont
        ElseIf Not pTheme.BaseFont Is Nothing Then
            Set lControl.Font = pTheme.BaseFont
        End If
    ElseIf TypeOf lControl Is CommandButton Or _
        TypeOf lControl Is Shape _
    Then
        ' nothing for these
    ElseIf TypeOf lControl Is DTPicker Then
        lControl.CalendarBackColor = pTheme.BackColor
        lControl.CalendarForeColor = pTheme.TextForeColor
        lControl.CalendarTitleBackColor = pTheme.TextBackColor
        lControl.CalendarTitleForeColor = pTheme.TextForeColor
        lControl.CalendarTrailingForeColor = AdjustColorIntensity(pTheme.TextForeColor, 0.5)
    ElseIf TypeOf lControl Is Object  Then
        On Error Resume Next
        If TypeOf lControl.object Is IThemeable Then
            If Err.Number = 0 Then
                On Error GoTo Err
                Dim lThemeable As IThemeable
                Set lThemeable = lControl.object
                lThemeable.Theme = pTheme
            Else
                On Error GoTo Err
            End If
        Else
            On Error GoTo Err
        End If
    End If
Next
        
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gFinishConfigChangeMonitoring()
Const ProcName As String = "gFinishConfigChangeMonitoring"
On Error GoTo Err

If Not mConfigChangeMonitor Is Nothing Then
    mConfigChangeMonitor.Finish
    Set mConfigChangeMonitor = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gGetSplashScreen() As Form
Const ProcName As String = "gGetSplashScreen"
On Error GoTo Err

If mSplash Is Nothing Then Set mSplash = New fSplash
mSplash.Show vbModeless
mSplash.Initialise
mSplash.Refresh
SetWindowLong mSplash.hWnd, GWL_EXSTYLE, GetWindowLong(mSplash.hWnd, GWL_EXSTYLE) Or WS_EX_TOPMOST
SetWindowPos mSplash.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

Set gGetSplashScreen = mSplash

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
                ByVal pAppInstanceConfig As ConfigurationSection) As Boolean
Const ProcName As String = "gLoadMainForm"
On Error GoTo Err

Do
    gGetSplashScreen
    
    Dim lMainForm As New fTradeSkilDemo
    
    LogMessage "Configuring TradeBuild with config: " & pAppInstanceConfig.InstanceQualifier
    Dim lTradeBuildAPI As TradeBuildAPI
    Set lTradeBuildAPI = configureTradeBuild(mConfigStore, pAppInstanceConfig)
    
    If lTradeBuildAPI Is Nothing Then
        LogMessage "Failed to configure TradeBuild with config: " & pAppInstanceConfig.InstanceQualifier
    End If
    
    LogMessage "Successfully configured TradeBuild with config: " & pAppInstanceConfig.InstanceQualifier
    LogMessage "Loading main form for config: " & pAppInstanceConfig.InstanceQualifier
    
    Dim lErrorMessage As String
    If lMainForm.Initialise(lTradeBuildAPI, mConfigStore, pAppInstanceConfig, lErrorMessage) Then Exit Do

    lTradeBuildAPI.ServiceProviders.RemoveAll
    gUnloadMainForm
    gUnloadSplashScreen
    
    Dim userResponse As Long
    userResponse = MsgBox( _
            "The configuration failed to operate correctly, for this " & vbCrLf & _
            "reason:" & vbCrLf & vbCrLf & _
            lErrorMessage & vbCrLf & vbCrLf & _
            "Would you like to manually correct the configuration?" & vbCrLf & vbCrLf & _
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
        Set pAppInstanceConfig = gShowConfigEditor(mConfigStore, pAppInstanceConfig, Nothing, pCentreWindow:=True)
    ElseIf userResponse = vbNo Then
        LogMessage "Creating a new default app instance configuration"
        Set pAppInstanceConfig = AddAppInstanceConfig(mConfigStore, _
                                                DefaultAppInstanceConfigName, _
                                                ConfigFlagIncludeDefaultBarFormatterLibrary Or _
                                                    ConfigFlagIncludeDefaultStudyLibrary Or _
                                                    ConfigFlagSetAsDefault)
    Else
        gLoadMainForm = False
        Exit Function
    End If
    
Loop

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
                ByVal pTheme As ITheme, _
                Optional ByVal title As String)
Const ProcName As String = "gModelessMsgBox"
On Error GoTo Err

ModelessMsgBox prompt, buttons, title, gMainForm, pTheme

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

Public Sub gSetFinished()
mFinished = True
End Sub

Public Function gShowConfigEditor( _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pCurrAppInstanceConfig As ConfigurationSection, _
                ByVal pTheme As ITheme, _
                Optional ByVal pParentForm As Form, _
                Optional ByVal pCentreWindow As Boolean = False) As ConfigurationSection
Const ProcName As String = "gShowConfigEditor"
On Error GoTo Err

If mConfigEditor Is Nothing Then
    Set mConfigEditor = New fConfigEditor
       
    mConfigEditor.Initialise pConfigStore, pCurrAppInstanceConfig, pCentreWindow
End If
If Not pTheme Is Nothing Then mConfigEditor.Theme = pTheme
mConfigEditor.Show vbModal, pParentForm
Set gShowConfigEditor = mConfigEditor.SelectedAppConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gUnloadConfigEditor()
If Not mConfigEditor Is Nothing Then
    Unload mConfigEditor
    Set mConfigEditor = Nothing
End If
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gUnloadSplashScreen()
Const ProcName As String = "gUnloadSplashScreen"
On Error GoTo Err

If mSplash Is Nothing Then Exit Sub
LogMessage "Unloading Splash Scren"
Unload mSplash
Set mSplash = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Main()
Const ProcName As String = "Main"
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)

If showCommandLineOptions() Then Exit Sub

Set mFatalErrorHandler = New FatalErrorHandler

ApplicationGroupName = "TradeWright"
ApplicationName = gAppTitle
SetupDefaultLogging Command

TaskQuantumMillisecs = 32

Set mConfigStore = getConfigStore
If mConfigStore Is Nothing Then
    LogMessage "Program exiting at user request"
    TerminateTWUtilities
    Exit Sub
End If

Set mConfigChangeMonitor = New ConfigChangeMonitor
mConfigChangeMonitor.Initialise mConfigStore

gGetSplashScreen
loadChartStyles mConfigStore
ensureBuiltInChartStylesExist
    
Dim lAppInstanceConfig As ConfigurationSection
Set lAppInstanceConfig = GetAppInstanceConfiguration(mConfigStore)
If lAppInstanceConfig Is Nothing Then
    LogMessage "Program exiting at user request"
    gFinishConfigChangeMonitoring
    gUnloadSplashScreen
    gUnloadConfigEditor
    TerminateTWUtilities
    Exit Sub
End If

If Not gLoadMainForm(lAppInstanceConfig) Then
    LogMessage "Program exiting at user request"
    gFinishConfigChangeMonitoring
    gUnloadSplashScreen
    gUnloadConfigEditor
    TerminateTWUtilities
    Exit Sub
End If

Do While Forms.Count > 0 And Not mFinished
    Wait 50
Loop

gUnloadSplashScreen
gFinishConfigChangeMonitoring

LogMessage "Application exiting"

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
gNotifyUnhandledError ProcName, ModuleName
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
    gUnloadSplashScreen
    
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
        Set pAppInstanceConfig = gShowConfigEditor(pConfigStore, pAppInstanceConfig, Nothing, pCentreWindow:=True)
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

Private Sub ensureBuiltInChartStylesExist()
Const ProcName As String = "ensureBuiltInChartStylesExist"
On Error GoTo Err

setupChartStyleAppDefault
setupChartStyleBlack
setupChartStyleDarkBlueFade
setupChartStyleGoldFade

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function GetAppInstanceConfiguration(ByVal pConfigStore As ConfigurationStore) As ConfigurationSection
Const ProcName As String = "GetAppInstanceConfiguration"
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
    Set lAppInstanceConfig = GetAppInstanceConfig(pConfigStore, configName)
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

Set GetAppInstanceConfiguration = lAppInstanceConfig

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
    LogMessage Err.Description, LogLevelSevere
    On Error GoTo Err
    If queryReplaceConfigFile Then Set lConfigStore = createNewConfigStore
ElseIf lConfigStore Is Nothing Then
    On Error GoTo Err
    LogMessage "The configuration file does not exist."
    If queryCreateNewConfigFile Then Set lConfigStore = createNewConfigStore
ElseIf Not IsValidConfigurationFile(lConfigStore) Then
    LogMessage "The configuration file is invalid."
    Set lConfigStore = Nothing
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

Private Sub setupChartStyleAppDefault()
Const ProcName As String = "setupChartStyleAppDefault"
On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameAppDefault) Then Exit Sub

Dim lCursorTextStyle As New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = &H80&
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0

Dim lFont As New StdFont
lFont.name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

ReDim GradientFillColors(1) As Long
GradientFillColors(0) = RGB(192, 192, 192)
GradientFillColors(1) = RGB(248, 248, 248)
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors

lDefaultRegionStyle.SessionEndGridLineStyle.Color = &HD0D0D0
lDefaultRegionStyle.SessionStartGridLineStyle.Color = &HD0D0D0
lDefaultRegionStyle.XGridLineStyle.Color = &HD0D0D0
lDefaultRegionStyle.YGridLineStyle.Color = &HD0D0D0
    
Dim lxAxisRegionStyle As ChartRegionStyle
Set lxAxisRegionStyle = GetDefaultChartXAxisRegionStyle.Clone
lxAxisRegionStyle.XCursorTextStyle = lCursorTextStyle
GradientFillColors(0) = RGB(230, 236, 207)
GradientFillColors(1) = RGB(222, 236, 215)
lxAxisRegionStyle.BackGradientFillColors = GradientFillColors
    
Dim lDefaultYAxisRegionStyle As ChartRegionStyle
Set lDefaultYAxisRegionStyle = GetDefaultChartYAxisRegionStyle.Clone
lDefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle
GradientFillColors(0) = RGB(234, 246, 254)
GradientFillColors(1) = RGB(226, 246, 255)
lDefaultYAxisRegionStyle.BackGradientFillColors = GradientFillColors
    
Dim lCrosshairLineStyle As New LineStyle
lCrosshairLineStyle.Color = &H7F

Dim lChartStyle As ChartStyle
Set lChartStyle = ChartStylesManager.Add(ChartStyleNameAppDefault, _
                        ChartStylesManager.DefaultStyle, _
                        lDefaultRegionStyle, _
                        lxAxisRegionStyle, _
                        lDefaultYAxisRegionStyle, _
                        lCrosshairLineStyle)

lChartStyle.ChartBackColor = RGB(192, 192, 192)
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupChartStyleBlack()
Const ProcName As String = "setupChartStyleBlack"
On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameBlack) Then Exit Sub

Dim lCursorTextStyle As New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = vbRed
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0

Dim lFont As StdFont
Set lFont = New StdFont
lFont.name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

ReDim GradientFillColors(1) As Long
GradientFillColors(0) = &H202020
GradientFillColors(1) = &H202020
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors

lDefaultRegionStyle.XGridLineStyle.Color = &H303030
lDefaultRegionStyle.YGridLineStyle.Color = &H303030
lDefaultRegionStyle.SessionEndGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle.Color = &H303030
lDefaultRegionStyle.SessionStartGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle.Color = &H303030
    
Dim lxAxisRegionStyle As ChartRegionStyle
Set lxAxisRegionStyle = GetDefaultChartXAxisRegionStyle.Clone
GradientFillColors(0) = RGB(0, 0, 0)
GradientFillColors(1) = RGB(0, 0, 0)
lxAxisRegionStyle.BackGradientFillColors = GradientFillColors
lxAxisRegionStyle.XCursorTextStyle = lCursorTextStyle

Dim lGridTextStyle As New TextStyle
lGridTextStyle.Box = True
lGridTextStyle.BoxFillWithBackgroundColor = True
lGridTextStyle.BoxStyle = LineInvisible
lGridTextStyle.Color = &HD0D0D0
lxAxisRegionStyle.XGridTextStyle = lGridTextStyle
    
Dim lDefaultYAxisRegionStyle As ChartRegionStyle
Set lDefaultYAxisRegionStyle = GetDefaultChartYAxisRegionStyle.Clone
GradientFillColors(0) = RGB(0, 0, 0)
GradientFillColors(1) = RGB(0, 0, 0)
lDefaultYAxisRegionStyle.BackGradientFillColors = GradientFillColors
lDefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle
lDefaultYAxisRegionStyle.YGridTextStyle = lGridTextStyle
    
Dim lCrosshairLineStyle As New LineStyle
lCrosshairLineStyle.Color = &H80&

Dim lChartStyle As ChartStyle
Set lChartStyle = ChartStylesManager.Add(ChartStyleNameBlack, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        lxAxisRegionStyle, _
                        lDefaultYAxisRegionStyle, _
                        lCrosshairLineStyle)

lChartStyle.ChartBackColor = &H202020

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupChartStyleDarkBlueFade()
Const ProcName As String = "setupChartStyleDarkBlueFade"
On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameDarkBlueFade) Then Exit Sub

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

ReDim GradientFillColors(1) As Long
GradientFillColors(0) = &H643232
GradientFillColors(1) = &H804040
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors
    
lDefaultRegionStyle.XGridLineStyle.Color = &H505050
lDefaultRegionStyle.YGridLineStyle.Color = &H505050
    
lDefaultRegionStyle.SessionEndGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle.Color = &H505050
    
lDefaultRegionStyle.SessionStartGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle.Color = &H505050

Dim lCrosshairLineStyle As New LineStyle
lCrosshairLineStyle.Color = vbRed

Dim lChartStyle As ChartStyle
Set lChartStyle = ChartStylesManager.Add(ChartStyleNameDarkBlueFade, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        , _
                        , _
                        lCrosshairLineStyle)
lChartStyle.ChartBackColor = &H643232

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupChartStyleGoldFade()
Const ProcName As String = "setupChartStyleGoldFade"
On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameGoldFade) Then Exit Sub

Dim lCursorTextStyle As New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = &H80&
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0

Dim lFont As New StdFont
lFont.name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone

ReDim GradientFillColors(1) As Long
GradientFillColors(0) = &H82DFE6
GradientFillColors(1) = &HEBFAFB
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors
    
lDefaultRegionStyle.XGridLineStyle.Color = &HE0E0E0
lDefaultRegionStyle.YGridLineStyle.Color = &HE0E0E0
    
lDefaultRegionStyle.SessionEndGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle.Color = &HE0E0E0
    
lDefaultRegionStyle.SessionStartGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle.Color = &HE0E0E0

Dim lCrosshairLineStyle As New LineStyle
lCrosshairLineStyle.Color = 127

Dim lChartStyle As ChartStyle
Set lChartStyle = ChartStylesManager.Add(ChartStyleNameGoldFade, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        , _
                        , _
                        lCrosshairLineStyle)
lChartStyle.ChartBackColor = &H82DFE6

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetWindowThemeOff(ByVal phWnd As Long)
Dim result As Long
result = SetWindowTheme(phWnd, vbNullString, "")
End Sub

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



