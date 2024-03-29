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

'================================================================================
' Interfaces
'================================================================================

Implements IChangeListener
Implements ICollectionChangeListener
Implements IDeferredAction


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

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTicker                                     As Ticker
Private WithEvents mTickerEv                        As Ticker
Attribute mTickerEv.VB_VarHelpID = -1

Private mStrategy                                   As IStrategy
Private mStrategyResourceContext                    As ResourceContext

Private mContract As IContract

Private mPositionManagementStrategyFactories        As EnumerableCollection
Private mPosnMgmtStrategyResourceContexts           As Collection

Private mDefaultParameters                          As Parameters
Private mParams                                     As Parameters

Private mInitialisationContext                      As InitialisationContext
Private mInitialisationCompleted                    As Boolean

Private mTradingContext                             As TradingContext

Private mTickers                                    As Tickers

Private mStartReplayWhenReady                       As Boolean

Private mBracketOrderNotificationRequests           As New BracketOrderNotifyReqs

Private mBracketOrderMapping                        As New EnumerableCollection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Randomize
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

If TypeOf ev.Source Is IBracketOrder Then
    Select Case ev.ChangeType
    Case BracketOrderChangeTypes.BracketOrderEntryOrderFilled, _
            BracketOrderChangeTypes.BracketOrderStopLossOrderChanged, _
            BracketOrderChangeTypes.BracketOrderCompleted
    Case Else
        Exit Sub
    End Select
    
    Dim lBracketOrder As IBracketOrder: Set lBracketOrder = ev.Source
    
    Dim en As Enumerator: Set en = mBracketOrderNotificationRequests.GetEnumerator(lBracketOrder)
    Dim lRequest As NotificationRequest
    Do While en.MoveNext
        Set lRequest = en.Current
        setStrategyRunner lRequest.EventSink, lRequest.ResourceContext
        If ev.ChangeType = BracketOrderChangeTypes.BracketOrderCompleted Then
            lBracketOrder.RemoveChangeListener Me
            lRequest.EventSink.NotifyBracketOrderCompletion GetResourceIdForBracketOrder(lBracketOrder)
        ElseIf ev.ChangeType = BracketOrderChangeTypes.BracketOrderEntryOrderFilled Then
            lRequest.EventSink.NotifyBracketOrderFill GetResourceIdForBracketOrder(lBracketOrder)
        ElseIf ev.ChangeType = BracketOrderChangeTypes.BracketOrderStopLossOrderChanged Then
            lRequest.EventSink.NotifyBracketOrderStopLossAdjusted GetResourceIdForBracketOrder(lBracketOrder)
        End If
        gClearStrategyRunner
    Loop
    
ElseIf TypeOf ev.Source Is PositionManager Then
    Dim lPositionManager As PositionManager
    Set lPositionManager = ev.Source
    If ev.ChangeType = PositionManagerChangeTypes.PositionSizeChanged Then
        If lPositionManager.PositionSize = 0 And _
            lPositionManager.PendingPositionSize = 0 Then
            setStrategyRunner mStrategy, mStrategyResourceContext
            If lPositionManager.IsSimulated Then
                mStrategy.NotifyNoSimulatedPositions
            Else
                mStrategy.NotifyNoLivePositions
            End If
            gClearStrategyRunner
        End If
    End If
ElseIf TypeOf ev.Source Is OrderContext Then
    Dim lOrderContext As OrderContext
    Set lOrderContext = ev.Source
    If lOrderContext.IsSimulated Then
        mStrategy.NotifyTradingReadinessChange
    Else
        mStrategy.NotifySimulatedTradingReadinessChange
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ICollectionChangeListener Interface Members
'@================================================================================

Private Sub ICollectionChangeListener_Change(ev As CollectionChangeEventData)
Const ProcName As String = "ICollectionChangeListener_Change"
On Error GoTo Err

If Not TypeOf ev.Source Is OrderContexts Then Exit Sub

If Not ev.ChangeType = CollItemAdded Then Exit Sub

Dim lOrderContext As OrderContext
Set lOrderContext = ev.AffectedItem
lOrderContext.AddChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Assert mTicker.State = MarketDataSourceStateReady, "MarketDataSource not ready"
getStrategyHost.TickerCreated mTicker

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mFutureWaiter Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then
ElseIf ev.Future.IsFaulted Then
    getStrategyHost.ContractInvalid ev.Future.ErrorMessage
ElseIf TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    
    Set mTicker = mTickers.CreateTicker(ev.Future, False, mContract.Specifier.Key)
    Set mTickerEv = mTicker
    
    If Not mStrategy Is Nothing Then setupLogging mContract.Specifier.LocalSymbol
ElseIf TypeOf ev.Future.Value Is BoxedValue Then
    Dim b As BoxedValue: Set b = ev.Future.Value
    gLogger.Log "Trading context says: " & b.Value, ModuleName, ProcName
    
    initialiseStrategy

    initialisePositionManagementStrategyFactories
    
    setStrategyRunner mStrategy, mStrategyResourceContext
    mStrategy.Start mTradingContext
    gClearStrategyRunner

    startReplayIfReady
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

Private Sub mReplayController_ReplayStarted()
Const ProcName As String = "mReplayController_ReplayStarted"
On Error GoTo Err

getStrategyHost.NotifyReplayStarted

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mReplayController_TickfileCompleted(ev As TickfileEventData, ByVal pEventsPlayed As Long)
Const ProcName As String = "mReplayController_TickfileCompleted"
On Error GoTo Err

getStrategyHost.NotifyTickfileCompleted ev.Specifier, pEventsPlayed

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTickerEv_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mTicker_StateChange"
On Error GoTo Err

Dim lState As MarketDataSourceStates
lState = ev.State

If lState <> MarketDataSourceStateReady Then Exit Sub

Set mTickerEv = Nothing
DeferAction Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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

Public Property Get DefaultParameters() As Parameters
Set DefaultParameters = mDefaultParameters
End Property

Public Property Get Parameters() As Parameters
Set Parameters = mParams
End Property

Friend Property Get PositionManagementStrategyFactories() As EnumerableCollection
Set PositionManagementStrategyFactories = mPositionManagementStrategyFactories
End Property

Friend Property Get Strategy() As IStrategy
Set Strategy = mStrategy
End Property

Friend Property Get Timestamp() As Date
Const ProcName As String = "Timestamp"
On Error GoTo Err

Timestamp = mTradingContext.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Function GetDefaultParameters( _
                ByVal pStrategy As IStrategy, _
                ByVal pPositionManagementStrategyFactories As EnumerableCollection) As Parameters
Const ProcName As String = "GetDefaultParameters"
On Error GoTo Err

Set mStrategy = pStrategy
If Not mContract Is Nothing Then setupLogging mContract.Specifier.LocalSymbol
Set mPositionManagementStrategyFactories = pPositionManagementStrategyFactories

Dim lStrategyParams As New Parameters
Set mDefaultParameters = lStrategyParams

If Not mStrategy Is Nothing Then
    setStrategyRunner mStrategy, Nothing
    mStrategy.DefineDefaultParameters
    gClearStrategyRunner
End If

Dim lPMFactory As IPosnMgmtStrtgyFactory
For Each lPMFactory In mPositionManagementStrategyFactories
    Dim lPMFactoryParams As New Parameters
    Set mDefaultParameters = lPMFactoryParams
    
    setStrategyRunner lPMFactory, Nothing
    lPMFactory.DefineDefaultParameters
    gClearStrategyRunner
    
    Set lStrategyParams = mergeParameters(lStrategyParams, lPMFactoryParams)
Next

Set mDefaultParameters = lStrategyParams
Set GetDefaultParameters = lStrategyParams

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetPositionManagementStrategyResourceContext(ByVal pFactory As IPosnMgmtStrtgyFactory) As ResourceContext
Const ProcName As String = "GetPositionManagementStrategyResourceContext"
On Error GoTo Err

Set GetPositionManagementStrategyResourceContext = mPosnMgmtStrategyResourceContexts(GetObjectKey(pFactory))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetResourceIdForBracketOrder( _
                ByVal pBracketOrder As IBracketOrder) As ResourceIdentifier
Const ProcName As String = "GetResourceIdForBracketOrder"
On Error GoTo Err

Set GetResourceIdForBracketOrder = mBracketOrderMapping.Item(GetObjectKey(pBracketOrder))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub InitialisationCompleted()
Const ProcName As String = "InitialisationCompleted"
On Error GoTo Err

gLogger.Log "Strategy: " & mStrategy.Name & ": Initialisation completed", ProcName, ModuleName

mInitialisationCompleted = True
mTradingContext.AllowUnprotectedPositions = mInitialisationContext.AllowUnprotectedPositions

Set mInitialisationContext = Nothing

mTicker.PositionManager.AddChangeListener Me
mTicker.PositionManager.OrderContexts.AddCollectionChangeListener Me

mTicker.PositionManagerSimulated.AddChangeListener Me
mTicker.PositionManagerSimulated.OrderContexts.AddCollectionChangeListener Me

startReplayIfReady

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

Friend Sub MapBracketOrderToResourceId( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pIdentifer As ResourceIdentifier)
Const ProcName As String = "MapBracketOrderToResourceId"
On Error GoTo Err

mBracketOrderMapping.Add pIdentifer, GetObjectKey(pBracketOrder)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PrepareSymbol( _
                ByVal pSymbol As IContractSpecifier)
Const ProcName As String = "PrepareSymbol"
On Error GoTo Err

Set mTickers = getStrategyHost.RealtimeTickers
mFutureWaiter.Add FetchContract(pSymbol, mContractStorePrimary, mContractStoreSecondary)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PrepareTickFile(ByVal pTickfiles As TickFileSpecifiers)
Const ProcName As String = "PrepareTickFile"
On Error GoTo Err

Assert mReplayController Is Nothing, "mReplayController is not Nothing "
    
Dim lTickfileDataManager As TickfileDataManager
Set lTickfileDataManager = CreateSequentialTickDataManager(pTickfiles, _
                                            mTickfileStoreInput, _
                                            mStudyLibraryManager, _
                                            mContractStorePrimary, _
                                            mContractStoreSecondary, _
                                            MarketDataSourceOptUseExchangeTimeZone, _
                                            pReplaySpeed:=0)

Set mReplayController = lTickfileDataManager.ReplayController

Dim lOrderManager As New OrderManager
Set mTickers = CreateTickers(lTickfileDataManager, mHistoricalDataStoreInput, lOrderManager, mOrderSubmitterFactoryLive, mOrderSubmitterFactorySimulated)

mFutureWaiter.Add mReplayController.TickStream(0).ContractFuture

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Friend Sub RequestBracketOrderNotification( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pStrategy As Object, _
                ByVal pResourceContext As ResourceContext)
Const ProcName As String = "RequestBracketOrderNotification"
On Error GoTo Err

mBracketOrderNotificationRequests.Add pBracketOrder, pStrategy, pResourceContext
pBracketOrder.AddChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetCurrent( _
                ByVal pStrategy As Object, _
                ByVal pResourceContext As ResourceContext)
Const ProcName As String = "SetCurrent"
On Error GoTo Err

If pResourceContext Is Nothing Then Set pResourceContext = mStrategyResourceContext
setStrategyRunner pStrategy, pResourceContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetNotCurrent()
gClearStrategyRunner
End Sub

Public Sub StartLiveData()
Const ProcName As String = "StartLiveData"
On Error GoTo Err

mTicker.StartMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartReplay()
Const ProcName As String = "StartReplay"
On Error GoTo Err

mStartReplayWhenReady = True
startReplayIfReady

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartStrategy( _
                ByVal pStrategy As IStrategy, _
                ByVal pParams As Parameters)
Const ProcName As String = "StartStrategy"
On Error GoTo Err

Set mStrategy = pStrategy
Set mParams = pParams

LogParameters mParams

StartLiveData
Set mTradingContext = New TradingContext
mFutureWaiter.Add mTradingContext.Initialise(Me, mStrategy, mTicker)

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

Private Sub initialisePositionManagementStrategyFactories()
Const ProcName As String = "initialisePositionManagementStrategyFactories"
On Error GoTo Err

Set mPosnMgmtStrategyResourceContexts = New Collection

Dim lPMFactory As IPosnMgmtStrtgyFactory
For Each lPMFactory In mPositionManagementStrategyFactories
    Dim lResourceContext As New ResourceContext
    mPosnMgmtStrategyResourceContexts.Add lResourceContext, GetObjectKey(lPMFactory)
    setStrategyRunner mPositionManagementStrategyFactories, lResourceContext
    lPMFactory.Initialise mInitialisationContext
    gClearStrategyRunner
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseStrategy()
Const ProcName As String = "initialiseStrategy"
On Error GoTo Err

mInitialisationCompleted = False
Set mInitialisationContext = New InitialisationContext
mInitialisationContext.Initialise getStrategyHost, Me, mTicker

Set mStrategyResourceContext = New ResourceContext
setStrategyRunner mStrategy, mStrategyResourceContext
mStrategy.Initialise mInitialisationContext
gClearStrategyRunner

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub LogParameters(ByVal pParams As Parameters)
Const ProcName As String = "logParameters"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Private Sub setStrategyRunner( _
                ByVal pStrategy As Object, _
                ByVal pResourceContext As ResourceContext)
Const ProcName As String = "setStrategyRunner"
On Error GoTo Err

gSetStrategyRunner Me, mInitialisationContext, mTradingContext, pResourceContext, pStrategy

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupLogging(ByVal pSymbol As String)
Const ProcName As String = "setupLogging"
On Error GoTo Err

Static sLoggingSetup As Boolean

If sLoggingSetup Then Exit Sub
sLoggingSetup = True

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
GetLogger("defaultstrategyhost").AddLogListener lLogfile
GetLogger("log").AddLogListener lLogfile
GetLogger("tradebuild.log.orderutils.contractresolution").AddLogListener lLogfile
GetLogger("position.order").AddLogListener lLogfile
GetLogger("position.simulatedorder").AddLogListener lLogfile

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

Private Sub startReplayIfReady()
Const ProcName As String = "startReplayIfReady"
On Error GoTo Err

If Not mStartReplayWhenReady Then Exit Sub

If mReplayController Is Nothing Then
    gLogger.Log "Not ready to start: " & mStrategy.Name & ": not got ReplayController", ProcName, ModuleName
    Exit Sub
End If

If Not mInitialisationCompleted Then
    gLogger.Log "Not ready to start: " & mStrategy.Name & ": initialisation not completed", ProcName, ModuleName
    Exit Sub
End If

mStartReplayWhenReady = False

LogMessage "Tickfile replay started"
mTicker.StartMarketData
Set mTickfileReplayTC = mReplayController.StartReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

