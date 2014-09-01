VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StrategyHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Implements StateChangeListener

'================================================================================
' Events
'================================================================================

Event ContractInvalid(ByVal pMessage As String)
Event ReplayCompleted()
Event ReplayProgress( _
                ByVal tickfileTimestamp As Date, _
                ByVal eventsPlayed As Long, _
                ByVal percentComplete As Single)
Event TickerCreated(ByVal pTicker As Ticker)

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                            As String = "StrategyHost"
'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTradeBuildAPI                              As TradeBuildAPI
Private WithEvents mReplayController                As ReplayController
Attribute mReplayController.VB_VarHelpID = -1
Private WithEvents mTickfileReplayTC                As TaskController
Attribute mTickfileReplayTC.VB_VarHelpID = -1

Private mHostController                             As New HostController

'Private mAccount As IAccount
Private mUseMoneyManagement                         As Boolean
Private mMoneyManager                               As IMoneyManager

Private mLogProfitProfile                           As Boolean
Private mLogDummyProfitProfile                      As Boolean
Private mResultsPath                                As String

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTicker                                     As Ticker

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Randomize
End Sub

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

If TypeOf ev.source Is Ticker Then
    If ev.State = MarketDataSourceStates.MarketDataSourceStateReady Then
        Set mTicker = ev.source
        RaiseEvent TickerCreated(mTicker)
        
        LogMessage "Tickfile replay started"
        Set mTickfileReplayTC = mReplayController.StartReplay
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mFutureWaiter Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Dim lContracts As IContracts
Set lContracts = ev.Future.Value

If lContracts.Count = 0 Then
    RaiseEvent ContractInvalid("No such contract")
ElseIf lContracts.Count > 1 Then
    RaiseEvent ContractInvalid("Contract not uniquely specified")
Else
    gTB.Tickers.CreateTicker CreateFuture(lContracts.ItemAtIndex(1)), False, Me
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mReplayController Event Handlers
'================================================================================

Private Sub mReplayController_ReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Long)
Const ProcName As String = "mReplayController_ReplayProgress"
On Error GoTo Err

RaiseEvent ReplayProgress(pTickfileTimestamp, pEventsPlayed, pPercentComplete)

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
RaiseEvent ReplayCompleted

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let LogDummyProfitProfile(ByVal val As Boolean)
mLogDummyProfitProfile = val
End Property

Friend Property Let LogProfitProfile(ByVal val As Boolean)
mLogProfitProfile = val
End Property

Friend Property Let ResultsPath(ByVal val As String)
mResultsPath = val
End Property

Friend Property Let UseMoneyManagement(ByVal val As Boolean)
mUseMoneyManagement = val
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub PlayTickFile(ByVal pTickfile As TickfileSpecifier)
Const ProcName As String = "PlayTickFile"
On Error GoTo Err

If mReplayController Is Nothing Then
    Dim lTickfiles As New tickfileSpecifiers
    lTickfiles.Add pTickfile
    
    Dim lTickfileDataManager As TickfileDataManager
    Set lTickfileDataManager = CreateTickDataManager(lTickfiles, _
                                                mTradeBuildAPI.TickfileStoreInput, _
                                                mTradeBuildAPI.StudyLibraryManager, _
                                                mTradeBuildAPI.ContractStorePrimary, _
                                                mTradeBuildAPI.ContractStoreSecondary, _
                                                MarketDataSourceOptUseExchangeTimeZone, _
                                                , _
                                                , _
                                                0, _
                                                250)
    
    Set mReplayController = lTickfileDataManager.ReplayController
    
    Dim lOrderManager As New OrderManager
    Dim lTickers As Tickers
    Set lTickers = CreateTickers(lTickfileDataManager, mTradeBuildAPI.StudyLibraryManager, mTradeBuildAPI.HistoricalDataStoreInput, lOrderManager, , mTradeBuildAPI.OrderSubmitterFactorySimulated)
    
    Dim lTicker As Ticker
    Set lTicker = lTickers.CreateTicker(mReplayController.TickStream(0).ContractFuture, False)
    lTicker.AddStateChangeListener Me
ElseIf mReplayController.ReplayInProgress Then
    LogMessage "Tickfile replay resumed"
    mReplayController.ResumeReplay
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Friend Function SetStrategy( _
                ByVal pStrategy As IStrategy, _
                ByVal pStopStrategyFactory As IPosnMgmtStrtgyFactory) As Parameters
Const ProcName As String = "SetStrategy"
On Error GoTo Err

Set mHostController = New HostController
Set SetStrategy = mHostController.Initialise(pStrategy, pStopStrategyFactory)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub StartTesting(ByVal pSymbol As String)
Const ProcName As String = "StartTesting"
On Error GoTo Err

prepare pSymbol

mFutureWaiter.Add gTB.ContractStorePrimary.FetchContracts(CreateContractSpecifier(pSymbol))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StopTesting()
Const ProcName As String = "StopTesting"
On Error GoTo Err

If mReplayController Is Nothing Then
    mTicker.Finish
Else
    mReplayController.StopReplay
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub prepare(ByVal pSymbol As String)
Const ProcName As String = "prepare"
On Error GoTo Err

If mUseMoneyManagement Then
'    Set mAccount = New cAccount
'    mAccount.Load "rlk"
'    Set mMoneyManager = New cMoneyManager
'    mMoneyManager.account = mAccount
'    mEngine.moneyManager = mMoneyManager
End If

If mResultsPath = "" Then mResultsPath = ApplicationSettingsFolder & "TestResults\"
If Right$(mResultsPath, 1) <> "\" Then mResultsPath = mResultsPath & "\"

Dim lFilenameDiscriminator As String
lFilenameDiscriminator = CStr(Int(1000000 * Rnd() + 1))

Dim lLogfile As FileLogListener
Set lLogfile = CreateFileLogListener(mResultsPath & "Logs\" & _
                                    mStrategy.Name & _
                                    "-" & _
                                    pSymbol & _
                                    "-" & lFilenameDiscriminator & ".log")
GetLogger("position.order").AddLogListener lLogfile
GetLogger("position.simulatedorder").AddLogListener lLogfile
GetLogger("position.drawdown").AddLogListener lLogfile
GetLogger("position.simulatedorder").AddLogListener lLogfile
GetLogger("strategy").AddLogListener lLogfile
GetLogger("position.moneymanagement").AddLogListener lLogfile

Set lLogfile = CreateFileLogListener(mResultsPath & "Orders\" & _
                                    mStrategy.Name & _
                                    "-" & _
                                    pSymbol & _
                                    "-" & lFilenameDiscriminator & ".log")
GetLogger("position.orderdetail").AddLogListener lLogfile

Set lLogfile = CreateFileLogListener(mResultsPath & "Orders\" & _
                                    mStrategy.Name & _
                                    "-" & _
                                    pSymbol & _
                                    "-" & lFilenameDiscriminator & _
                                    "-Profile" & ".log")
GetLogger("position.bracketorderprofile").AddLogListener lLogfile

Set lLogfile = CreateFileLogListener(mResultsPath & "Orders\" & _
                                    mStrategy.Name & _
                                    "-" & _
                                    pSymbol & _
                                    "-" & lFilenameDiscriminator & _
                                    "-TradeProfile" & ".log")
GetLogger("position.tradeprofile").AddLogListener lLogfile

If mLogProfitProfile Then
    Set lLogfile = CreateFileLogListener(mResultsPath & "Orders\" & _
                                        mStrategy.Name & _
                                        "-" & _
                                        pSymbol & _
                                        "-" & lFilenameDiscriminator & _
                                        "-Profit" & ".log")
    GetLogger("position.profitprofile").AddLogListener lLogfile
End If

Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub



