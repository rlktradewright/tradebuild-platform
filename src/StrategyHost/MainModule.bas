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

Private mStrategyRunner                             As StrategyRunner

Private mFatalErrorHandler                          As FatalErrorHandler

Public gTB                                          As TradeBuildAPI

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

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(Command)

If lClp.Switch("?") Then
    MsgBox vbCrLf & getUsageString, , "Usage"
    Exit Sub
End If

ApplicationGroupName = "TradeWright"
ApplicationName = "StrategyHost"
SetupDefaultLogging Command

Dim lNoUI As Boolean
If lClp.Switch("noui") Then lNoUI = True

Dim lRun As Boolean
If lClp.Switch("run") Then lRun = True

Dim lLiveTrades As Boolean
If lClp.Switch("livetrades") Then lLiveTrades = True

Dim lSymbol As String
lSymbol = lClp.Arg(0)
If lSymbol = "" And lNoUI Then
    LogMessage "No symbol supplied"
    If Not lNoUI And lRun Then MsgBox "Error - no symbol argument supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

Dim lStrategyClassName As String
lStrategyClassName = lClp.Arg(1)
If lStrategyClassName = "" And lNoUI Then
    LogMessage "No strategy supplied"
    If Not lNoUI And lRun Then MsgBox "Error - no strategy class name argument supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

Dim lStopStrategyFactoryClassName As String
lStopStrategyFactoryClassName = lClp.Arg(2)
If lStopStrategyFactoryClassName = "" And lNoUI Then
    LogMessage "No stop strategy factory supplied"
    If Not lNoUI And lRun Then MsgBox "Error - no stop strategy factory class name argument supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

Dim lPermittedSPRoles As ServiceProviderRoles
lPermittedSPRoles = SPRoleContractDataPrimary + _
                    SPRoleHistoricalDataInput + _
                    SPRoleOrderSubmissionLive + _
                    SPRoleOrderSubmissionSimulated

If Not lLiveTrades And Not lNoUI Then lPermittedSPRoles = lPermittedSPRoles + SPRoleTickfileInput

If lClp.Switch("tws") Then lPermittedSPRoles = lPermittedSPRoles + SPRoleRealtimeData

Set gTB = CreateTradeBuildAPI(, lPermittedSPRoles)

If lClp.Switch("tws") Then
    If Not setupTwsServiceProviders(lClp.switchValue("tws"), lLiveTrades) Then
        MsgBox "Error setting up tws service provider - see log at " & DefaultLogFileName(Command) & vbCrLf & getUsageString, vbCritical, "Error"
        Exit Sub
    End If
End If

If lClp.Switch("db") Then
    If Not setupDbServiceProviders(lClp.switchValue("db"), Not (lLiveTrades Or lNoUI)) Then
        MsgBox "Error setting up database service providers - see log at " & DefaultLogFileName(Command) & vbCrLf & getUsageString, vbCritical, "Error"
        Exit Sub
    End If
Else
    MsgBox "/db not supplied: " & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

If Not setupSimulateOrderServiceProviders(lLiveTrades) Then
    MsgBox "Error setting up simulated orders service provider(s) - see log at " & DefaultLogFileName(Command) & vbCrLf & getUsageString, vbCritical, "Error"
    Exit Sub
End If

If Not gTB.StartServiceProviders Then
    MsgBox "One or more service providers failed to start: see logfile"
    Exit Sub
End If

gTB.StudyLibraryManager.AddBuiltInStudyLibrary

Dim lUseMoneyManagement As Boolean
If lClp.Switch("umm") Or _
    lClp.Switch("UseMoneyManagement") _
Then
    lUseMoneyManagement = True
End If

Dim lResultsPath As String
If lClp.Switch("ResultsPath") Then
    lResultsPath = lClp.switchValue("ResultsPath")
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
    
    If lClp.Switch("tws") Then
        mForm.SymbolText.Enabled = True
        mForm.SymbolText.Text = lSymbol
    End If
    mForm.ResultsPathText = lResultsPath
    mForm.NoMoneyManagement = IIf(lUseMoneyManagement, 0, 1)
    mForm.StrategyCombo.Text = lStrategyClassName
    mForm.StopStrategyFactoryCombo.Text = lStopStrategyFactoryClassName
    
    mForm.Show vbModeless
    
    If lRun Then
        mForm.Start
    End If

    Do While Forms.Count > 0
        Wait 50
    Loop
    
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
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getUsageString() As String
getUsageString = _
            "strategyhost  [symbol]" & vbCrLf & _
            "              [strategy class name]" & vbCrLf & _
            "              [stop strategy factory class name]" & vbCrLf & _
            "              [/tws:[Server],[Port],[ClientId]" & vbCrLf & _
            "              [/db:[server],[servertype],[database]" & vbCrLf & _
            "              [/livetrades]" & vbCrLf & _
            "              [/logpath:path]" & vbCrLf & _
            "              [/noUI]" & vbCrLf & _
            "              [/resultsPath:path]" & vbCrLf & _
            "              [/run]" & vbCrLf & _
            "              [/useMoneyManagement  |  /umm]"
End Function

Private Function setupDbServiceProviders( _
                ByVal switchValue As String, _
                ByVal pAllowTickfiles As Boolean) As Boolean
Dim clp As CommandLineParser
Dim Server As String
Dim dbtypeStr As String
Dim dbtype As DatabaseTypes
Dim database As String
Dim username As String
Dim password As String

Dim failpoint As Long
On Error GoTo Err

Set clp = CreateCommandLineParser(switchValue, ",")

setupDbServiceProviders = True

On Error Resume Next
Server = clp.Arg(0)
dbtypeStr = clp.Arg(1)
database = clp.Arg(2)
username = clp.Arg(3)
password = clp.Arg(4)
On Error GoTo Err

dbtype = DatabaseTypeFromString(dbtypeStr)
If dbtype = DbNone Then
    LogMessage "Error: invalid dbtype"
    setupDbServiceProviders = False
End If

If username <> "" And password = "" Then
    LogMessage "Password not supplied"
    setupDbServiceProviders = False
End If
    
If setupDbServiceProviders Then
    gTB.ServiceProviders.Add _
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

    gTB.ServiceProviders.Add _
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
        gTB.ServiceProviders.Add _
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

Private Function setupSimulateOrderServiceProviders(ByVal pLiveTrades As Boolean) As Boolean
On Error GoTo Err

If Not pLiveTrades Then
    gTB.ServiceProviders.Add _
                        ProgId:="TradeBuild27.OrderSimulatorSP", _
                        Enabled:=True, _
                        Name:="TradeBuild Exchange Simulator for Main Orders", _
                        ParamString:="Role=LIVE", _
                        Description:="Simulated order submission for main orders"
End If

gTB.ServiceProviders.Add _
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

Private Function setupTwsServiceProviders( _
                ByVal switchValue As String, _
                ByVal pAllowLiveTrades As Boolean) As Boolean
On Error GoTo Err

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

setupTwsServiceProviders = True

On Error Resume Next
Dim Server As String
Server = clp.Arg(0)

Dim Port As String
Port = clp.Arg(1)

Dim ClientId As String
ClientId = clp.Arg(2)
On Error GoTo Err

If Port = "" Then
    Port = "7496"
ElseIf Not IsInteger(Port, 1) Then
        LogMessage "Error: Port must be a positive integer > 0"
        setupTwsServiceProviders = False
End If
    
If ClientId = "" Then
    ClientId = "1215339864"
ElseIf Not IsInteger(ClientId, 0) Then
        LogMessage "Error: ClientId must be an integer >= 0"
        setupTwsServiceProviders = False
End If
    
If setupTwsServiceProviders Then
    gTB.ServiceProviders.Add _
                        ProgId:="IBTWSSP27.RealtimeDataServiceProvider", _
                        Enabled:=True, _
                        ParamString:="Role=PRIMARY" & _
                                    ";Server=" & Server & _
                                    ";Port=" & Port & _
                                    ";Client Id=" & ClientId & _
                                    ";Provider Key=IB;Keep Connection=True", _
                        Description:="Realtime data"
    
    If pAllowLiveTrades Then
        gTB.ServiceProviders.Add _
                            ProgId:="IBTWSSP27.OrderSubmissionSrvcProvider", _
                            Enabled:=True, _
                            ParamString:="Server=" & Server & _
                                        ";Port=" & Port & _
                                        ";Client Id=" & ClientId & _
                                        ";Provider Key=IB;Keep Connection=True", _
                            Description:="Live order submission"
    End If
End If

Exit Function

Err:
LogMessage Err.Description, LogLevelSevere
setupTwsServiceProviders = False
End Function



