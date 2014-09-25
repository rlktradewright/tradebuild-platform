VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StrategyRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                            As String = "StrategyRunner"
'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mStrategyHostRef                            As WeakReference

Private mTickfileStoreInput                         As ITickfileStore
Private mStudyLibraryManager                        As StudyLibraryManager
Private mContractStorePrimary                       As IContractStore
Private mContractStoreSecondary                     As IContractStore
Private mHistoricalDataStoreInput                   As IHistoricalDataStore
Private mOrderSubmitterFactoryLive                  As IOrderSubmitterFactory
Private mOrderSubmitterFactorySimulated             As IOrderSubmitterFactory

Private WithEvents mReplayController                As ReplayController
Attribute mReplayController.VB_VarHelpID = -1
Private WithEvents mTickfileReplayTC                As TaskController
Attribute mTickfileReplayTC.VB_VarHelpID = -1

'Private mAccount As IAccount

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTicker                                     As Ticker

Private mStrategy                                   As IStrategy
Private mStopStrategyFactory                        As IPosnMgmtStrtgyFactory

Private mParams                                     As Parameters

Private mInitialisationContext                      As New InitialisationContext

Private mTickers                                    As Tickers
Private mIsTickfileReplay                           As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Randomize
End Sub

'================================================================================
' mFutureWaiter Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    getStrategyHost.ContractInvalid ev.Future.ErrorMessage
ElseIf ev.Future.IsAvailable Then
    Dim lContract As IContract
    Set lContract = ev.Future.Value
    
    Set mTicker = mTickers.CreateTicker(ev.Future, False, lContract.Specifier.Key)
    Assert mTicker.State = MarketDataSourceStateReady
    mTicker.StartMarketData

    setupLogging lContract.Specifier.LocalSymbol
    getStrategyHost.TickerCreated mTicker
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mReplayController Event Handlers
'================================================================================

Private Sub mReplayController_NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "mReplayController_NotifyEvent"
On Error GoTo Err

Set ev.Source = Me
getStrategyHost.NotifyReplayEvent ev

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mReplayController_ReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Long)
Const ProcName As String = "mReplayController_ReplayProgress"
On Error GoTo Err

getStrategyHost.NotifyReplayProgress pTickfileTimestamp, pEventsPlayed, pPercentComplete

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mTickfileReplayTC Event Handlers
'================================================================================

Private Sub mTickfileReplayTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTickfileReplayTC_Completed"
On Error GoTo Err

Set mReplayController = Nothing
getStrategyHost.NotifyReplayCompleted

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub InitialisationContextIsReady(ByVal pContext As InitialisationContext)
Const ProcName As String = "InitialisationContextIsReady"
On Error GoTo Err

mStrategy.Initialise pContext, mParams
mStopStrategyFactory.Initialise pContext, mParams

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialisationCompleted()
Const ProcName As String = "InitialisationCompleted"
On Error GoTo Err

Dim lTradingContext As New TradingContext
lTradingContext.Initialise mStrategy, mTicker
mStrategy.Start lTradingContext, mStopStrategyFactory

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pStrategyHost As IStrategyHost)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStrategyHostRef = CreateWeakReference(pStrategyHost)

Set mTickfileStoreInput = getStrategyHost.TickfileStoreInput
Set mStudyLibraryManager = getStrategyHost.StudyLibraryManager
Set mContractStorePrimary = getStrategyHost.ContractStorePrimary
Set mContractStoreSecondary = getStrategyHost.ContractStoreSecondary
Set mHistoricalDataStoreInput = getStrategyHost.HistoricalDataStoreInput
Set mOrderSubmitterFactoryLive = getStrategyHost.OrderSubmitterFactoryLive
Set mOrderSubmitterFactorySimulated = getStrategyHost.OrderSubmitterFactorySimulated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PrepareSymbol( _
                ByVal pSymbol As String)
Const ProcName As String = "PrepareSymbol"
On Error GoTo Err

mIsTickfileReplay = False
Set mTickers = getStrategyHost.RealtimeTickers
mFutureWaiter.Add FetchContract(CreateContractSpecifier(pSymbol), mContractStorePrimary, mContractStoreSecondary)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PrepareTickFile(ByVal pTickfile As TickfileSpecifier)
Const ProcName As String = "PrepareTickFile"
On Error GoTo Err

Assert mReplayController Is Nothing, "mReplayController is not Nothing "
    
mIsTickfileReplay = False

Dim lTickfiles As New TickFileSpecifiers
lTickfiles.Add pTickfile

Dim lTickfileDataManager As TickfileDataManager
Set lTickfileDataManager = CreateTickDataManager(lTickfiles, _
                                            mTickfileStoreInput, _
                                            mStudyLibraryManager, _
                                            mContractStorePrimary, _
                                            mContractStoreSecondary, _
                                            MarketDataSourceOptUseExchangeTimeZone, _
                                            pReplaySpeed:=0)

Set mReplayController = lTickfileDataManager.ReplayController

Dim lOrderManager As New OrderManager
Set mTickers = CreateTickers(lTickfileDataManager, mStudyLibraryManager, mHistoricalDataStoreInput, lOrderManager, mOrderSubmitterFactoryLive, mOrderSubmitterFactorySimulated)

mFutureWaiter.Add mReplayController.TickStream(0).ContractFuture

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Public Function SetStrategy( _
                ByVal pStrategy As IStrategy, _
                ByVal pStopStrategyFactory As IPosnMgmtStrtgyFactory) As Parameters
Const ProcName As String = "SetStrategy"
On Error GoTo Err

Set mStrategy = pStrategy
Set mStopStrategyFactory = pStopStrategyFactory
Set SetStrategy = mergeParameters(pStrategy.DefaultParameters, pStopStrategyFactory.DefaultParameters)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub StartReplay()
Const ProcName As String = "StartReplay"
On Error GoTo Err

If Not mReplayController Is Nothing Then
    LogMessage "Tickfile replay started"
    Set mTickfileReplayTC = mReplayController.StartReplay
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartStrategy( _
                ByVal pParams As Parameters)
Const ProcName As String = "StartStrategy"
On Error GoTo Err

Set mParams = pParams

If getStrategyHost.LogParameters Then
    Dim s As String
    s = "Strategy Parameters:" & vbCrLf
    If Not pParams Is Nothing Then
        Dim lParam As Parameter
        For Each lParam In pParams
            s = s & vbTab & lParam.Name & "=" & lParam.Value & vbCrLf
        Next
    End If
    gStrategyLogger.Log LogLevelNormal, s
End If

mInitialisationContext.Initialise getStrategyHost, Me, mTicker


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopTesting()
Const ProcName As String = "StopTesting"
On Error GoTo Err

If Not mReplayController Is Nothing Then
    ' prevent event handler being fired on completion, which would
    ' reload the main form again
    Set mTickfileReplayTC = Nothing
    If mReplayController.ReplayInProgress Then mReplayController.StopReplay
    Set mReplayController = Nothing
End If
    
If Not mTicker Is Nothing Then mTicker.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function getStrategyHost() As IStrategyHost
Set getStrategyHost = mStrategyHostRef.Target
End Function

Private Function mergeParameters(ByVal pParams1 As Parameters, ByVal pParams2 As Parameters) As Parameters
Const ProcName As String = "mergeParameters"
On Error GoTo Err

Const DefaultValue As String = "***$Default$***"

Dim lParams As New Parameters

Dim lParam As Parameter
For Each lParam In pParams1
    lParams.SetParameterValue lParam.Name, lParam.Value
Next

For Each lParam In pParams2
    If lParams.GetParameterValue(lParam.Name, DefaultValue) = DefaultValue Then lParams.SetParameterValue lParam.Name, lParam.Value
Next

Set mergeParameters = lParams

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setupLogging(ByVal pSymbol As String)
Const ProcName As String = "setupLogging"
On Error GoTo Err

If getStrategyHost.UseMoneyManagement Then
'    Set mAccount = New cAccount
'    mAccount.Load "rlk"
'    Set mMoneyManager = New cMoneyManager
'    mMoneyManager.account = mAccount
'    mEngine.moneyManager = mMoneyManager
End If

Dim lResultsPath As String
lResultsPath = getStrategyHost.ResultsPath
If lResultsPath = "" Then lResultsPath = ApplicationSettingsFolder & "\TestResults\"
If Right$(lResultsPath, 1) <> "\" Then lResultsPath = lResultsPath & "\"

Dim lFilenameDiscriminator As String
lFilenameDiscriminator = CStr(Int(1000000 * Rnd() + 1))

Dim lLogfile As FileLogListener
Set lLogfile = CreateFileLogListener(lResultsPath & "Logs\" & _
                                        mStrategy.Name & _
                                        "-" & _
                                        pSymbol & _
                                        "-" & lFilenameDiscriminator & ".log", _
                                    includeTimestamp:=False, _
                                    includeLogLevel:=False)
GetLogger("position.order").AddLogListener lLogfile
GetLogger("position.simulatedorder").AddLogListener lLogfile
GetLogger("strategy").AddLogListener lLogfile
GetLogger("position.moneymanagement").AddLogListener lLogfile

Set lLogfile = CreateFileLogListener(lResultsPath & "Orders\" & _
                                        mStrategy.Name & _
                                        "-" & _
                                        pSymbol & _
                                        "-" & lFilenameDiscriminator & _
                                        ".log", _
                                    includeTimestamp:=False, _
                                    includeLogLevel:=False)
GetLogger("position.orderdetail").AddLogListener lLogfile

Set lLogfile = CreateFileLogListener(lResultsPath & "Orders\" & _
                                        mStrategy.Name & _
                                        "-" & _
                                        pSymbol & _
                                        "-" & lFilenameDiscriminator & _
                                        "-Profile" & ".log", _
                                    includeTimestamp:=False, _
                                    includeLogLevel:=False)
GetLogger("position.bracketorderprofilestring").AddLogListener lLogfile

Set lLogfile = CreateFileLogListener(lResultsPath & "Orders\" & _
                                        mStrategy.Name & _
                                        "-" & _
                                        pSymbol & _
                                        "-" & lFilenameDiscriminator & _
                                        "-TradeProfile" & ".log", _
                                    includeTimestamp:=False, _
                                    includeLogLevel:=False)
GetLogger("position.tradeprofile").AddLogListener lLogfile

If getStrategyHost.LogProfitProfile Then
    Set lLogfile = CreateFileLogListener(lResultsPath & "Orders\" & _
                                        mStrategy.Name & _
                                        "-" & _
                                        pSymbol & _
                                        "-" & lFilenameDiscriminator & _
                                        "-Profit" & ".log", _
                                    includeTimestamp:=False, _
                                    includeLogLevel:=False)
    GetLogger("position.profitprofile").AddLogListener lLogfile
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

