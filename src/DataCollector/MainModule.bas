Attribute VB_Name = "MainModule"
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                                As String = "DataCollector27"
Public Const AppName                                    As String = "TradeBuild Data Collector"

Private Const ModuleName                                As String = "MainModule"

Public Const AttributeNameBidAskBars                    As String = "WriteBidAndAskBars"
Public Const AttributeNameEnabled                       As String = "Enabled"
Public Const AttributeNameIncludeMktDepth               As String = "IncludeMarketDepth"

Public Const ConfigSectionCollectionControl             As String = "CollectionControl"
Public Const ConfigSectionContract                      As String = "Contract"
Public Const ConfigSectionContracts                     As String = "Contracts"
Public Const ConfigSectionContractspecifier             As String = "ContractSpecifier"
Public Const ConfigSectionTickdata                      As String = "TickData"

Public Const ConfigSettingContractSpecCurrency          As String = ConfigSectionContractspecifier & "&Currency"
Public Const ConfigSettingContractSpecExpiry            As String = ConfigSectionContractspecifier & "&Expiry"
Public Const ConfigSettingContractSpecExchange          As String = ConfigSectionContractspecifier & "&Exchange"
Public Const ConfigSettingContractSpecLocalSymbol       As String = ConfigSectionContractspecifier & "&LocalSymbol"
Public Const ConfigSettingContractSpecMultiplier        As String = ConfigSectionContractspecifier & "&Multiplier"
Public Const ConfigSettingContractSpecRight             As String = ConfigSectionContractspecifier & "&Right"
Public Const ConfigSettingContractSpecSecType           As String = ConfigSectionContractspecifier & "&SecType"
Public Const ConfigSettingContractSpecStrikePrice       As String = ConfigSectionContractspecifier & "&StrikePrice"
Public Const ConfigSettingContractSpecSymbol            As String = ConfigSectionContractspecifier & "&Symbol"
Public Const ConfigSettingContractSpecTradingClass      As String = ConfigSectionContractspecifier & "&TradingClass"

Public Const ConfigFileVersion                          As String = "1.0"

Public Const ConfigNodeContractSpecs                    As String = "Contract Specifications"
Public Const ConfigNodeServiceProviders                 As String = "Service Providers"
Public Const ConfigNodeParameters                       As String = "Parameters"

Public Const ConfigSettingWriteBarData                  As String = ConfigSectionCollectionControl & "&WriteBarData"
Public Const ConfigSettingWriteTickData                 As String = ConfigSectionCollectionControl & "&WriteTickData"
Public Const ConfigSettingWriteTickDataFormat           As String = ConfigSectionTickdata & "&Format"
Public Const ConfigSettingWriteTickDataPath             As String = ConfigSectionTickdata & "&Path"

Private Const DefaultAppInstanceConfigName              As String = "Default Config"

Public Const PermittedSPRoles                           As Long = ServiceProviderRoles.SPRoleRealtimeData Or _
                                                                    ServiceProviderRoles.SPRoleContractDataPrimary Or _
                                                                    ServiceProviderRoles.SPRoleHistoricalDataOutput Or _
                                                                    ServiceProviderRoles.SPRoleTickfileOutput


' command line switch indicating which configuration to load
' when the programs starts (if not specified, the default configuration
' is loaded)
Public Const SwitchConfig                               As String = "config"

Public Const SwitchSetup                                As String = "setup"

Public Const SwitchConcurrency                          As String = "concurrency"
Public Const SwitchQuantum                              As String = "quantum"

Public Const SwitchLogBackup                            As String = "logbackup"
Public Const SwitchLogOverwrite                         As String = "logoverwrite"

Public Const TWSClientId                                As Long = -1
Public Const TWSConnectRetryInterval                    As Long = 10
Public Const TWSPort                                    As Long = 7496
Public Const TWSServer                                  As String = ""

Public Const TickfilesPath                              As String = "C:\Data\Tickfiles"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================


'@================================================================================
' Member variables
'@================================================================================

Private mIsInDev                                        As Boolean

Public gStop                                            As Boolean

Private mCLParser                                       As CommandLineParser

Private mConfig                                         As ConfigurationSection

Private mNoAutoStart                                    As Boolean
Private mNoUI                                           As Boolean
Private mLeftOffset                                     As Long
Private mRightOffset                                    As Long
Private mPosX                                           As Single
Private mPosY                                           As Single

Private mDataCollector                                  As DataCollector

Private mStartTimeDescriptor                            As String
Private mEndTimeDescriptor                              As String
Private mExitTimeDescriptor                             As String

Private mNoDataRestartSecs                              As Long

Private mConfigManager                                  As ConfigManager

Private mFatalErrorHandler                              As FatalErrorHandler

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

Public Property Get AppTitle() As String
AppTitle = AppName & _
                " v" & _
                App.Major & "." & App.Minor
End Property

Public Property Get configFilename() As String
Const ProcName As String = "configFilename"
On Error GoTo Err

Dim fn As String
fn = mCLParser.Arg(0)
If fn = "" Then
    fn = ApplicationSettingsFolder & "\settings.xml"
End If
configFilename = fn

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Sub gHandleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

' we don't do the following since if the program is running unattended we may want to definitely
' end the program at this point so that it can be restarted. What we really need is an alerter
' service provider that can provide a notification in some form that a fault has occurred.
'If Not mNoUI Then
'    MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
'            "Please email the log file located at" & vbCrLf & vbCrLf & _
'            "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
'            "to support@tradewright.com", _
'            vbCritical, _
'            "Fatal error"
'End If

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly. Note that normally one would use the
' End statement to terminate a VB6 program abruptly. However the TWUtilities component interferes
' with the End statement's processing and prevents proper shutdown, so we use the
' TWUtilities component's EndProcess method instead. (However if we are running in the
' development environment, then we call End because the EndProcess method kills the
' entire development environment as well which can have undesirable side effects if other
' components are also loaded.)

'If mIsInDev Then
'    ' this tells TWUtilities that we've now handled this unhandled error. Not actually
'    ' needed here because the End statement will prevent return to TWUtilities
'    UnhandledErrorHandler.Handled = True
'    End
'Else
'    EndProcess
'End If

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
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.description)
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
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub Main()
Const ProcName As String = "Main"
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler

Set mCLParser = CreateCommandLineParser(Command, " ")

If showHelp Then
    TerminateTWUtilities
    Exit Sub
End If

ApplicationGroupName = "TradeWright"
ApplicationName = AppTitle

Dim lLogOverwrite As Boolean
lLogOverwrite = mCLParser.Switch(SwitchLogOverwrite)

Dim lLogBackup As Boolean
lLogBackup = mCLParser.Switch(SwitchLogBackup)

LogMessage "Logfile is: " & SetupDefaultLogging(Command, lLogOverwrite, lLogBackup)
LogMessage "Loglevel is: " & LogLevelToString(DefaultLogLevel)

RunTasksAtLowerThreadPriority = False

mLeftOffset = -1
mRightOffset = -1

setTaskParameters

mNoUI = getNoUi

If getConfigFile Is Nothing Then
    createConfigFile configFilename, getConfigName
    If Not getConfig Then
        TerminateTWUtilities
        Exit Sub
    End If
    showConfig
ElseIf Not getConfig Then
    TerminateTWUtilities
    Exit Sub
ElseIf setup Then
    Exit Sub
End If

If Not configure Then
    TerminateTWUtilities
    Exit Sub
End If

mStartTimeDescriptor = getStartTimeDescriptor
mEndTimeDescriptor = getEndTimeDescriptor
mExitTimeDescriptor = getExitTimeDescriptor

mNoAutoStart = getNoAutostart

mNoDataRestartSecs = getNoDataRestartSecs

If mNoUI Then
    
    LogMessage "Creating data collector object"
    Set mDataCollector = CreateDataCollector(mConfigManager.ConfigurationFile, _
                                            mConfig.InstanceQualifier, _
                                            mStartTimeDescriptor, _
                                            mEndTimeDescriptor, _
                                            mExitTimeDescriptor)
    
    If mStartTimeDescriptor = "" Then
        LogMessage "Starting data collection"
        mDataCollector.StartCollection
    End If
    
    Do While Not gStop
        Wait 1000
    Loop
    
    LogMessage "Data Collector program exiting"
    
    TerminateTWUtilities
    
Else
    LogMessage "Creating data collector object"
    Dim f As New fDataCollectorUI
    Set mDataCollector = CreateDataCollector(mConfigManager.ConfigurationFile, _
                                            mConfig.InstanceQualifier, _
                                            IIf(mNoAutoStart, "", mStartTimeDescriptor), _
                                            mEndTimeDescriptor, _
                                            mExitTimeDescriptor, _
                                            5, _
                                            f, _
                                            f)
    
    LogMessage "Showing form"
    showMainForm f
End If


Exit Sub
Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function configure() As Boolean
Const ProcName As String = "configure"
On Error GoTo Err

If getConfigToLoad() Is Nothing Then
    notifyError "No configuration is available"
Else
    Set mConfig = getConfigToLoad
    LogMessage "Configuration in use: " & mConfig.InstanceQualifier
    configure = True
End If

Exit Function

Err:
configure = False
End Function

Public Sub createConfigFile(ByVal pConfigFilename As String, ByVal pConfigName As String)
Const ProcName As String = "createConfigFile"
On Error GoTo Err

LogMessage "No configuration file exists - creating skeleton configuration file"

Dim lBaseConfigFile As IConfigStoreProvider
Set lBaseConfigFile = CreateXMLConfigurationProvider(App.ProductName, ConfigFileVersion)

Dim lConfigStore As ConfigurationStore
Set lConfigStore = CreateConfigurationStore(lBaseConfigFile, pConfigFilename)
InitialiseConfigFile lConfigStore
Dim lAppConfig As ConfigurationSection
Set lAppConfig = AddAppInstanceConfig(lConfigStore, _
                    IIf(pConfigName <> "", pConfigName, DefaultAppInstanceConfigName), _
                    ConfigFlagSetAsDefault, _
                    PermittedSPRoles, _
                    pTWSServer:=TWSServer, _
                    pTWSPort:=TWSPort, _
                    pTwsClientId:=TWSClientId, _
                    pTwsConnectionRetryIntervalSecs:=TWSConnectRetryInterval, _
                    pTickfilesPath:=TickfilesPath)
lAppConfig.addConfigurationSection ConfigSectionCollectionControl
lAppConfig.addConfigurationSection ConfigSectionContracts

lConfigStore.Save pConfigFilename

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getConfig() As Boolean
Const ProcName As String = "getConfig"
On Error GoTo Err

Set mConfigManager = New ConfigManager
If mConfigManager.Initialise(configFilename) Then
    logConfigFileDetails
    getConfig = True
Else
    notifyError "The configuration file (" & _
                    configFilename & _
                    ") is not the correct format for this program"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getConfigFile() As IConfigStoreProvider
Const ProcName As String = "getConfigFile"
On Error GoTo Err

Dim lFso As New FileSystemObject
If lFso.FileExists(configFilename) Then Set getConfigFile = LoadConfigProviderFromXMLFile(configFilename)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getConfigName() As String
Const ProcName As String = "getConfigName"
On Error GoTo Err

If mCLParser.Switch(SwitchConfig) Then
    getConfigName = mCLParser.SwitchValue(SwitchConfig)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getConfigToLoad() As ConfigurationSection
Const ProcName As String = "getConfigToLoad"
On Error GoTo Err

Static configToLoad As ConfigurationSection

If configToLoad Is Nothing Then
    On Error Resume Next
    Set configToLoad = getNamedConfig()
    If Err.Number <> 0 Then Exit Function
    On Error GoTo Err

End If

Set getConfigToLoad = configToLoad

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function getEndTimeDescriptor() As String
Const ProcName As String = "getEndTimeDescriptor"
On Error GoTo Err

If mCLParser.Switch("endAt") Then
    getEndTimeDescriptor = mCLParser.SwitchValue("endAt")
End If
LogMessage "End at: " & getEndTimeDescriptor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getExitTimeDescriptor() As String
Const ProcName As String = "getExitTimeDescriptor"
On Error GoTo Err

If mCLParser.Switch("exitAt") Then
    getExitTimeDescriptor = mCLParser.SwitchValue("exitAt")
End If
LogMessage "Exit at: " & getExitTimeDescriptor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNamedConfig() As ConfigurationSection
Const ProcName As String = "getNamedConfig"
On Error GoTo Err

If getConfigName <> "" Then
    Set getNamedConfig = mConfigManager.AppConfig(getConfigName)
    If getNamedConfig Is Nothing Then
        notifyError "The required configuration does not exist: " & getConfigName
        Err.Raise ErrorCodes.ErrIllegalArgumentException
    End If
Else
    Set getNamedConfig = mConfigManager.DefaultAppConfig
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNoAutostart() As Boolean
Const ProcName As String = "getNoAutostart"
On Error GoTo Err

If mCLParser.Switch("noAutoStart") Then
    getNoAutostart = True
End If
LogMessage "Auto start: " & Not getNoAutostart

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNoDataRestartSecs() As Long
Const ProcName As String = "getNoDataRestartSecs"
On Error GoTo Err

If mCLParser.Switch("noDataRestartSecs") Then
    getNoDataRestartSecs = mCLParser.SwitchValue("noDataRestartSecs")
End If
LogMessage "No data restart interval (secs): " & getNoDataRestartSecs

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNoUi() As Boolean
Const ProcName As String = "getNoUi"
On Error GoTo Err

If mCLParser.Switch("noui") Then
    getNoUi = True
End If
LogMessage "Run with UI: " & Not getNoUi

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getStartTimeDescriptor() As String
Const ProcName As String = "getStartTimeDescriptor"
On Error GoTo Err

If mCLParser.Switch("startAt") Then
    getStartTimeDescriptor = mCLParser.SwitchValue("startAt")
End If
LogMessage "Start at: " & getStartTimeDescriptor

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

Private Sub logConfigFileDetails()
Const ProcName As String = "logConfigFileDetails"
On Error GoTo Err

LogMessage "Configuration file: " & configFilename

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyError( _
                ByVal message As String)
Const ProcName As String = "notifyError"
On Error GoTo Err

LogMessage message, LogLevels.LogLevelSevere
If Not mNoUI Then MsgBox message & vbCrLf & vbCrLf & "The program will close.", vbCritical, "Attention!"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setTaskParameters()
Const ProcName As String = "setTaskParameters"
On Error GoTo Err

RunTasksAtLowerThreadPriority = False
TaskConcurrency = 20
TaskQuantumMillisecs = 20

If mCLParser.Switch(SwitchConcurrency) Then
    If Not IsInteger(mCLParser.SwitchValue(SwitchConcurrency), 2) Then
        LogMessage "Argument '" & SwitchConcurrency & ":" & mCLParser.SwitchValue(SwitchConcurrency) & "' is invalid and has been ignored"
    Else
        TaskConcurrency = CLng(mCLParser.SwitchValue(SwitchConcurrency))
    End If
End If
LogMessage "Task concurrency=" & TaskConcurrency

If mCLParser.Switch(SwitchQuantum) Then
    If Not IsInteger(mCLParser.SwitchValue(SwitchQuantum), 1) Then
        LogMessage "Argument '" & SwitchQuantum & ":" & mCLParser.SwitchValue(SwitchQuantum) & "' is invalid and has been ignored"
    Else
        TaskQuantumMillisecs = CLng(mCLParser.SwitchValue(SwitchQuantum))
    End If
End If
LogMessage "Task quantum (millisecs)=" & TaskQuantumMillisecs

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function setup() As Boolean
Const ProcName As String = "setup"
On Error GoTo Err

If Not mCLParser.Switch(SwitchSetup) Then Exit Function
showConfig
setup = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SetWindowThemeOff(ByVal phWnd As Long)
Dim result As Long
result = SetWindowTheme(phWnd, vbNullString, "")
End Sub

Private Sub showConfig()
Const ProcName As String = "showConfig"
On Error GoTo Err

Dim f As fConfig
Set f = New fConfig
f.Initialise mConfigManager, False
f.Show vbModeless

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function showHelp() As Boolean
Const ProcName As String = "showHelp"
On Error GoTo Err

If Trim$(Command) <> "/?" And Trim$(Command) <> "-?" Then Exit Function
    
Dim s As String
s = vbCrLf & _
        "datacollector27 [configfilename]" & vbCrLf & _
        "                /setup " & vbCrLf & _
        "   or " & vbCrLf & _
        "datacollector27 [configfilename] " & vbCrLf & _
        "                [/config:configtoload] " & vbCrLf & _
        "                [/log:filename] " & vbCrLf & _
        "                [/posn:offsetfromleft,offsetfromtop]" & vbCrLf & _
        "                [/noAutoStart" & vbCrLf & _
        "                [/noUI]" & vbCrLf & _
        "                [/showMonitor]" & vbCrLf & _
        "                [/exitAt:[day]hh:mm]" & vbCrLf & _
        "                [/startAt:[day]hh:mm]" & vbCrLf & _
        "                [/endAt:[day]hh:mm]" & vbCrLf & _
        "                [/loglevel:levelName]" & vbCrLf
s = s & _
        "                [/noDataRestartSecs:n]" & vbCrLf & _
        "                [/concurrency:n]" & vbCrLf & _
        "                [/quantum:n]" & vbCrLf
s = s & vbCrLf & _
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
        "       All     or A"
s = s & vbCrLf & _
        "Example 1:" & vbCrLf & _
        "   datacollector27 /setup" & vbCrLf & _
        "       runs the data collector configurer, which enables you to define " & vbCrLf & _
        "       various configurations for use with different data collector " & vbCrLf & _
        "       instances. The default configuration file is used to store this" & vbCrLf & _
        "       information." & vbCrLf & _
        "Example 2:" & vbCrLf & _
        "   datacollector27 mysettings.xml /config:""US Futures""" & vbCrLf & _
        "       runs the data collector in accordance with the configuration" & vbCrLf & _
        "       called ""US Futures"" defined in the mysettings.xml file."
MsgBox s, , "Usage"
showHelp = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub showMainForm(ByVal pForm As fDataCollectorUI)
Const ProcName As String = "showMainForm"
On Error GoTo Err

If mCLParser.Switch("posn") Then
    Dim posnValue As String
    posnValue = mCLParser.SwitchValue("posn")
    
    If InStr(1, posnValue, ",") = 0 Then
        MsgBox "Error - posn value must be 'n,m'"
        Exit Sub
    End If
    
    If Not IsNumeric(Left$(posnValue, InStr(1, posnValue, ",") - 1)) Then
        MsgBox "Error - offset from left is not numeric"
        Exit Sub
    End If
    
    mPosX = Left$(posnValue, InStr(1, posnValue, ",") - 1)
    
    If Not IsNumeric(Right$(posnValue, Len(posnValue) - InStr(1, posnValue, ","))) Then
        MsgBox "Error - offset from top is not numeric"
        Exit Sub
    End If
    
    mPosY = Right$(posnValue, Len(posnValue) - InStr(1, posnValue, ","))
Else
    mPosX = Int(Int(Screen.Width / pForm.Width) * Rnd)
    mPosY = Int(Int(Screen.Height / pForm.Height) * Rnd)
End If

LogMessage "Form position: " & mPosX & "," & mPosY

pForm.Initialise mDataCollector, _
                mConfigManager, _
                getConfigName, _
                getNoAutostart, _
                CBool(mCLParser.Switch("showMonitor")), _
                mNoDataRestartSecs

pForm.Left = mPosX * pForm.Width
pForm.Top = mPosY * pForm.Height

pForm.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



