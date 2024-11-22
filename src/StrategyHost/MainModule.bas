Attribute VB_Name = "MainModule"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MainModule"

'@================================================================================
' Member variables
'@================================================================================

Private mForm                                       As fStrategyHost

Private mFatalErrorHandler                          As FatalErrorHandler

Private mTB                                         As TradeBuildAPI

Private mModel                                      As IStrategyHostModel

Private mStrategyHost                               As DefaultStrategyHost

Public gFinished                                    As Boolean

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

'@================================================================================
' Methods
'@================================================================================

Public Sub Main()
Const ProcName As String = "Main"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(Command)

If lClp.Switch("?") Then
    MsgBox vbCrLf & getUsageString, , "Usage"
    Exit Sub
End If

init lClp


Dim lNoUI As Boolean
If lClp.Switch("noui") Then lNoUI = True

Dim lRun As Boolean
If lClp.Switch("run") Then lRun = True

mModel.UseLiveBroker = Not lClp.Switch("simulateorders")

Dim lContractSpecString As String
lContractSpecString = lClp.switchValue("contract")
If lContractSpecString <> "" Then
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = CreateContractSpecifierFromString(lContractSpecString)
    If lContractSpec Is Nothing Then
        LogMessage "Invalid contract spec"
        If Not lNoUI And lRun Then MsgBox "Error - invalid contract supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
        Exit Sub
    End If
    mModel.Symbol = lContractSpec
ElseIf lNoUI Then
    LogMessage "No contract supplied"
    Exit Sub
ElseIf Not lNoUI And lRun Then MsgBox "Error - no contract supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

Dim lStrategyProgIds As String
lStrategyProgIds = lClp.switchValue("strategy")
If lStrategyProgIds = "" And lNoUI Then
    LogMessage "No strategy supplied"
    If Not lNoUI And lRun Then MsgBox "Error - no strategy class name argument supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

Dim lStopLossStrategyFactoryProgIds As String
lStopLossStrategyFactoryProgIds = lClp.switchValue("stoplossstrategy")
If lStopLossStrategyFactoryProgIds = "" And lNoUI Then
    LogMessage "No stop-loss strategy factory supplied"
    If Not lNoUI And lRun Then MsgBox "Error - no stop-loss strategy factory class name argument supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

Dim lTargetStrategyFactoryProgIds As String
lTargetStrategyFactoryProgIds = lClp.switchValue("targetstrategy")
If lTargetStrategyFactoryProgIds = "" And lNoUI Then
    LogMessage "No target strategy factory supplied"
    If Not lNoUI And lRun Then MsgBox "Error - no target strategy factory class name argument supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

If Not setupServiceProviders(lClp, Not mModel.UseLiveBroker, lNoUI) Then Exit Sub

If lClp.Switch("umm") Or _
    lClp.Switch("UseMoneyManagement") _
Then
    mModel.UseMoneyManagement = True
End If

If lClp.Switch("ResultsDir") Then
    mModel.ResultsPath = lClp.switchValue("ResultsDir")
End If

If lNoUI Then
'    Set mStrategyRunner = CreateStrategyRunner(Me)
'    mStrategyRunner.UseMoneyManagement = lUseMoneyManagement
'    mStrategyRunner.ResultsPath = lResultsPath
'    mStrategyRunner.SetStrategy CreateObject(lStrategyClassName), Nothing
'    mStrategyRunner.PrepareSymbol lSymbol
'    Set mStrategyRunner = Nothing
Else
    Set mForm = New fStrategyHost
    mForm.Show vbModeless
    
    Dim lFrameWidth As Long
    lFrameWidth = GetSystemMetrics(SM_CXSIZEFRAME) * Screen.TwipsPerPixelX
    
    Dim lFrameHeight As Long
    lFrameHeight = GetSystemMetrics(SM_CYSIZEFRAME) * Screen.TwipsPerPixelY
    
    mForm.Left = -lFrameWidth
    mForm.Top = -lFrameHeight
    
    Dim workarea As GDI_RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, VarPtr(workarea), 0
    
    mForm.Height = workarea.Bottom * Screen.TwipsPerPixelY + 2 * lFrameHeight
    mForm.Width = workarea.Right * Screen.TwipsPerPixelX / 2 + 2 * lFrameWidth
    
    Dim failpoint As String
    
    failpoint = "Dim lController As New DefaultStrategyHostCtlr"
    Dim lController As New DefaultStrategyHostCtlr
    
    failpoint = "Set mStrategyHost = New DefaultStrategyHost"
    Set mStrategyHost = New DefaultStrategyHost
    
    failpoint = "mStrategyHost.Initialise mModel, mForm, lController"
    mStrategyHost.Initialise mModel, mForm, lController
    
    failpoint = "mForm.Initialise mModel, lController"
    mForm.Initialise mModel, lController, lStrategyProgIds, lStopLossStrategyFactoryProgIds, lTargetStrategyFactoryProgIds
    
    If lRun Then
        mForm.Start
    End If

    Do While Not gFinished
        Wait 50
    Loop
    
    Set mForm = Nothing
    
    LogMessage "Removing all service providers"
    mTB.ServiceProviders.RemoveAll
    
    LogMessage "Application exiting"
    
    TerminateTWUtilities
End If

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
gNotifyUnhandledError ProcName, ModuleName, failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getUsageString() As String
getUsageString = _
            "strategyhost  [/contract:contractspec)]" & vbCrLf & _
            "              [/strategy:progId [; progid]...]" & vbCrLf & _
            "              [/stopstrategy:progId [; progid]...]]" & vbCrLf & _
            "              [/targetstrategy:progId [; progid]...]]" & vbCrLf & _
            "              [/tws:[server],[port],[clientId]" & vbCrLf & _
            "              [/db:[server],[servertype],[database]" & vbCrLf & _
            "              [/livetrades]" & vbCrLf & _
            "              [/log:path]" & vbCrLf & _
            "              [/loglevel:levelName]" & vbCrLf & _
            "              [/logoverwrite]" & vbCrLf & _
            "              [/logbackup]" & vbCrLf & _
            "              [/noUI]" & vbCrLf & _
            "              [/resultsPath:path]" & vbCrLf & _
            "              [/run]" & vbCrLf & _
            "              [/useMoneyManagement  |  /umm]" & vbCrLf & _
            vbCrLf & _
            " where" & vbCrLf & _
            vbCrLf
getUsageString = getUsageString & _
            "   levelname is one of:" & vbCrLf & _
            "       None    or 0" & vbCrLf & _
            "       Severe  or S" & vbCrLf & _
            "       Warning or W" & vbCrLf & _
            "       Info    or I" & vbCrLf & _
            "       Normal  or N" & vbCrLf & _
            "       Detail  or D" & vbCrLf & _
            "       Medium  or M" & vbCrLf & _
            "       High    or H" & vbCrLf & _
            "       All     or A"
End Function

Private Sub init(ByVal pClp As CommandLineParser)
Const ProcName As String = "init"
On Error GoTo Err

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler

ApplicationGroupName = "TradeWright"
ApplicationName = "StrategyHost"

Dim lLogOverwrite As Boolean
lLogOverwrite = pClp.Switch("logoverwrite")

Dim lLogBackup As Boolean
lLogBackup = pClp.Switch("logbackup")

SetupDefaultLogging Command, lLogOverwrite, lLogBackup

logProgramId

Set mModel = New DefaultStrategyHostModl
mModel.LogParameters = True
mModel.ShowChart = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logProgramId()
Const ProcName As String = "logProgramId"
On Error GoTo Err

Dim s As String
s = App.ProductName & _
    " V" & _
    App.Major & _
    "." & App.Minor & _
    "." & App.Revision & _
    IIf(App.FileDescription <> "", "-" & App.FileDescription, "") & _
    vbCrLf & _
    App.LegalCopyright
s = s & vbCrLf & "Arguments: " & Command
LogMessage s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function parseSymbol( _
                ByVal pSymbol As String) As IContractSpecifier
Const ProcName As String = "parseSymbol"
On Error GoTo Err

Const CurrencySwitch                         As String = "CURRENCY"
Const CurrencySwitch1                        As String = "CURR"
Const ExchangeSwitch                         As String = "EXCHANGE"
Const ExchangeSwitch1                        As String = "EXCH"
Const ExpirySwitch                           As String = "EXPIRY"
Const ExpirySwitch1                          As String = "EXP"
Const LocalSymbolSwitch                      As String = "LOCALSYMBOL"
Const LocalSymbolSwitch1                     As String = "LOCAL"
Const MultiplierSwitch                       As String = "MULTIPLIER"
Const MultiplierSwitch1                      As String = "MULT"
Const RightSwitch                            As String = "RIGHT"
Const SecTypeSwitch                          As String = "SECTYPE"
Const SecTypeSwitch1                         As String = "SEC"
Const SymbolSwitch                           As String = "SYMBOL"
Const SymbolSwitch1                          As String = "SYMB"
Const StrikeSwitch                           As String = "STRIKE"
Const StrikeSwitch1                          As String = "STR"
Const TradingClassSwitch                     As String = "TRADINGCLASS"

If Not Left$(pSymbol, 1) = "(" Or Not Right$(pSymbol, 1) = ")" Then Exit Function

pSymbol = Mid$(pSymbol, 2, Len(pSymbol) - 2)

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pSymbol, ";")

Dim validParams As Boolean
validParams = True

Dim lSectype As String: lSectype = lClp.switchValue(SecTypeSwitch)
If lSectype = "" Then lSectype = lClp.switchValue(SecTypeSwitch1)

Dim lExchange As String: lExchange = lClp.switchValue(ExchangeSwitch)
If lExchange = "" Then lExchange = lClp.switchValue(ExchangeSwitch1)

Dim lLocalSymbol As String: lLocalSymbol = lClp.switchValue(LocalSymbolSwitch)
If lLocalSymbol = "" Then lLocalSymbol = lClp.switchValue(LocalSymbolSwitch1)

Dim lSymbol As String: lSymbol = lClp.switchValue(SymbolSwitch)
If lSymbol = "" Then lSymbol = lClp.switchValue(SymbolSwitch1)

Dim lTradingClass As String: lTradingClass = lClp.switchValue(TradingClassSwitch)

Dim lCurrency As String: lCurrency = lClp.switchValue(CurrencySwitch)
If lCurrency = "" Then lCurrency = lClp.switchValue(CurrencySwitch1)

Dim lExpiry As String: lExpiry = lClp.switchValue(ExpirySwitch)
If lExpiry = "" Then lExpiry = lClp.switchValue(ExpirySwitch1)

Dim lMultiplier As String: lMultiplier = lClp.switchValue(MultiplierSwitch)
If lMultiplier = "" Then lMultiplier = lClp.switchValue(MultiplierSwitch1)
If lMultiplier = "" Then lMultiplier = "0.0"

Dim lStrike As String: lStrike = lClp.switchValue(StrikeSwitch)
If lStrike = "" Then lStrike = lClp.switchValue(StrikeSwitch1)
If lStrike = "" Then lStrike = "0.0"

Dim lRight As String: lRight = lClp.switchValue(RightSwitch)

Set parseSymbol = CreateContractSpecifier(lLocalSymbol, _
                                        lSymbol, _
                                        lTradingClass, _
                                        lExchange, _
                                        SecTypeFromString(lSectype), _
                                        lCurrency, _
                                        lExpiry, _
                                        CDbl(lMultiplier), _
                                        CDbl(lStrike), _
                                        OptionRightFromString(lRight))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupServiceProviders( _
                ByVal pClp As CommandLineParser, _
                ByVal pSimulateOrders As Boolean, _
                ByVal pNoUI As Boolean) As Boolean
Const ProcName As String = "setupServiceProviders"
On Error GoTo Err

Dim lPermittedSPRoles As ServiceProviderRoles
lPermittedSPRoles = SPRoleContractDataPrimary + _
                    SPRoleHistoricalDataInput + _
                    SPRoleOrderSubmissionLive + _
                    SPRoleOrderSubmissionSimulated

If pSimulateOrders And Not pNoUI Then lPermittedSPRoles = lPermittedSPRoles + SPRoleTickfileInput

If pClp.Switch("tws") Then lPermittedSPRoles = lPermittedSPRoles + SPRoleRealtimeData

Set mTB = CreateTradeBuildAPI(, lPermittedSPRoles)

If Not (pClp.Switch("tws") Or pClp.Switch("db")) Then
    MsgBox "Neither /tws nor /db supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Function
End If

Dim lOrdersViaPaperTWS As Boolean
If pClp.Switch("papertws") Then
    lOrdersViaPaperTWS = True
    If Not setupPaperTwsServiceProviders(pClp.switchValue("papertws")) Then
        MsgBox "Error setting up papertws service provider - see log at " & DefaultLogFileName(Command) & vbCrLf & getUsageString, vbCritical, "Error"
        Exit Function
    End If
End If

If pClp.Switch("tws") Then
    If Not setupTwsServiceProviders(pClp.switchValue("tws"), Not pClp.Switch("db"), Not lOrdersViaPaperTWS And Not pSimulateOrders) Then
        MsgBox "Error setting up tws service provider - see log at " & DefaultLogFileName(Command) & vbCrLf & getUsageString, vbCritical, "Error"
        Exit Function
    End If
ElseIf Not pSimulateOrders Then
    MsgBox "Must use simulated orders if /tws not supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Function
End If

If pClp.Switch("db") Then
    If Not setupDbServiceProviders(pClp.switchValue("db"), pSimulateOrders And Not pNoUI) Then
        MsgBox "Error setting up database service providers - see log at " & DefaultLogFileName(Command) & vbCrLf & getUsageString, vbCritical, "Error"
        Exit Function
    End If
End If

If Not setupSimulateOrderServiceProviders(pSimulateOrders) Then
    MsgBox "Error setting up simulated orders service provider(s) - see log at " & DefaultLogFileName(Command) & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Function
End If

If Not mTB.StartServiceProviders Then
    MsgBox "One or more service providers failed to start: see logfile"
    Exit Function
End If

mTB.StudyLibraryManager.AddBuiltInStudyLibrary

mModel.ContractStorePrimary = mTB.ContractStorePrimary
mModel.ContractStoreSecondary = mTB.ContractStoreSecondary
mModel.HistoricalDataStoreInput = mTB.HistoricalDataStoreInput
mModel.OrderSubmitterFactoryLive = mTB.OrderSubmitterFactoryLive
mModel.OrderSubmitterFactorySimulated = mTB.OrderSubmitterFactorySimulated
mModel.RealtimeTickers = mTB.Tickers
mModel.StudyLibraryManager = mTB.StudyLibraryManager
mModel.TickfileStoreInput = mTB.TickfileStoreInput

setupServiceProviders = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupDbServiceProviders( _
                ByVal switchValue As String, _
                ByVal pAllowTickfiles As Boolean) As Boolean

Dim failpoint As Long
On Error GoTo Err

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

setupDbServiceProviders = True

On Error Resume Next
Dim Server As String: Server = clp.Arg(0)

Dim dbtypeStr As String: dbtypeStr = clp.Arg(1)

Dim database As String: database = clp.Arg(2)

Dim username As String: username = clp.Arg(3)

Dim password As String: password = clp.Arg(4)
On Error GoTo Err

Dim dbtype As DatabaseTypes: dbtype = DatabaseTypeFromString(dbtypeStr)
If dbtype = DbNone Then
    LogMessage "Error: invalid dbtype"
    setupDbServiceProviders = False
End If

If username <> "" And password = "" Then
    LogMessage "Password not supplied"
    setupDbServiceProviders = False
End If
    
If setupDbServiceProviders Then
    mTB.ServiceProviders.Add _
                        ProgId:="TBInfoBase27.ContractInfoSrvcProvider", _
                        Enabled:=True, _
                        ParamString:="Role=PRIMARY" & _
                                    ";Database Name=" & database & _
                                    ";Database Type=" & dbtypeStr & _
                                    ";Server=" & Server & _
                                    ";User Name=" & username & _
                                    ";Password=" & password & _
                                    ";Use Synchronous Reads=True", _
                        Description:="Primary contract data"

    mTB.ServiceProviders.Add _
                        ProgId:="TBInfoBase27.HistDataServiceProvider", _
                        Enabled:=True, _
                        ParamString:="Role=INPUT" & _
                                    ";Database Name=" & database & _
                                    ";Database Type=" & dbtypeStr & _
                                    ";Server=" & Server & _
                                    ";User Name=" & username & _
                                    ";Password=" & password & _
                                    ";Use Synchronous Reads=False", _
                        Description:="Historical data input"

    If pAllowTickfiles Then
        mTB.ServiceProviders.Add _
                            ProgId:="TBInfoBase27.TickfileServiceProvider", _
                            Enabled:=True, _
                            ParamString:="Role=INPUT" & _
                                        ";Database Name=" & database & _
                                        ";Database Type=" & dbtypeStr & _
                                        ";Server=" & Server & _
                                        ";User Name=" & username & _
                                        ";Password=" & password & _
                                        ";Use Synchronous Reads=false", _
                            Description:="Tickfile input"
    End If
End If

Exit Function

Err:
LogMessage Err.Description, LogLevelSevere
setupDbServiceProviders = False
End Function

Private Function setupSimulateOrderServiceProviders(ByVal pSimulateOrders As Boolean) As Boolean
On Error GoTo Err

If pSimulateOrders Then
    mTB.ServiceProviders.Add _
                        ProgId:="TradeBuild27.OrderSimulatorSP", _
                        Enabled:=True, _
                        Name:="TradeBuild Exchange Simulator for Main Orders", _
                        ParamString:="Role=LIVE", _
                        Description:="Simulated order submission for main orders"
End If

mTB.ServiceProviders.Add _
                    ProgId:="TradeBuild27.OrderSimulatorSP", _
                    Enabled:=True, _
                    Name:="TradeBuild Exchange Simulator for Dummy Orders", _
                    ParamString:="Role=SIMULATED", _
                    Description:="Simulated order submission for dummy orders"

setupSimulateOrderServiceProviders = True
Exit Function

Err:
LogMessage Err.Description, LogLevelSevere
setupSimulateOrderServiceProviders = False
End Function

Private Function setupPaperTwsServiceProviders( _
                ByVal switchValue As String) As Boolean
On Error GoTo Err

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

setupPaperTwsServiceProviders = True

Dim Server As String
Server = clp.Arg(0)

Dim Port As String
Port = clp.Arg(1)
If Port = "" Then
    Port = "7497"
ElseIf Not IsInteger(Port, 1) Then
        LogMessage "Error: Port must be a positive integer > 0"
        setupPaperTwsServiceProviders = False
End If
    
Dim ClientId As String
ClientId = clp.Arg(2)
If ClientId = "" Then
    ClientId = "322255712"
ElseIf Not IsInteger(ClientId, 0) Then
        LogMessage "Error: ClientId must be an integer >= 0 and <= 999999999"
        setupPaperTwsServiceProviders = False
End If
    
If setupPaperTwsServiceProviders Then
    mTB.ServiceProviders.Add _
                        ProgId:="IBTWSSP27.OrderSubmissionSrvcProvider", _
                        Enabled:=True, _
                        ParamString:="Server=" & Server & _
                                    ";Port=" & Port & _
                                    ";Client Id=" & ClientId & _
                                    ";Provider Key=IB;Keep Connection=True", _
                        Description:="Paper-trading order submission"
End If

Exit Function

Err:
LogMessage Err.Description, LogLevelSevere
setupPaperTwsServiceProviders = False
End Function

Private Function setupTwsServiceProviders( _
                ByVal switchValue As String, _
                ByVal pNoDB As Boolean, _
                ByVal pAllowLiveTrades As Boolean) As Boolean
On Error GoTo Err

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

setupTwsServiceProviders = True

Dim Server As String
Server = clp.Arg(0)

Dim Port As String
Port = clp.Arg(1)
If Port = "" Then
    Port = "7496"
ElseIf Not IsInteger(Port, 1) Then
        LogMessage "Error: Port must be a positive integer > 0"
        setupTwsServiceProviders = False
End If
    
Dim ClientId As String
ClientId = clp.Arg(2)
If ClientId = "" Then
    ClientId = "215339864"
ElseIf Not IsInteger(ClientId, 0) Then
        LogMessage "Error: ClientId must be an integer >= 0 and <= 999999999"
        setupTwsServiceProviders = False
End If
    
If setupTwsServiceProviders Then
    mTB.ServiceProviders.Add _
                        ProgId:="IBTWSSP27.RealtimeDataServiceProvider", _
                        Enabled:=True, _
                        ParamString:="Server=" & Server & _
                                    ";Port=" & Port & _
                                    ";Client Id=" & ClientId & _
                                    ";Provider Key=IB;Keep Connection=True", _
                        Description:="Realtime data"
    
    If pAllowLiveTrades Then
        mTB.ServiceProviders.Add _
                            ProgId:="IBTWSSP27.OrderSubmissionSrvcProvider", _
                            Enabled:=True, _
                            ParamString:="Server=" & Server & _
                                        ";Port=" & Port & _
                                        ";Client Id=" & ClientId & _
                                        ";Provider Key=IB;Keep Connection=True", _
                            Description:="Live order submission"
    End If
    If pNoDB Then
        mTB.ServiceProviders.Add _
                            ProgId:="IBTWSSP27.ContractInfoServiceProvider", _
                            Enabled:=True, _
                            ParamString:="Role=PRIMARY" & _
                                        ";Server=" & Server & _
                                        ";Port=" & Port & _
                                        ";Client Id=" & ClientId & _
                                        ";Provider Key=IB;Keep Connection=True", _
                            Description:="Primary contract data"
    
        mTB.ServiceProviders.Add _
                            ProgId:="IBTWSSP27.HistDataServiceProvider", _
                            Enabled:=True, _
                            ParamString:="Server=" & Server & _
                                        ";Port=" & Port & _
                                        ";Client Id=" & ClientId & _
                                        ";Provider Key=IB;Keep Connection=True", _
                            Description:="Historical bar data retrieval"
    End If
End If

Exit Function

Err:
LogMessage Err.Description, LogLevelSevere
setupTwsServiceProviders = False
End Function



