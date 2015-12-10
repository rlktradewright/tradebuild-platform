VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements IContractFetchListener
Implements ITickfileReader

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================


Private Const ModuleName As String = "TickfileReader"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                  As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1

Private mCapabilities               As Long

Private mParent                     As TickfileStore
Private mReplayMonitor              As ITickfileReplayMonitor
Private mDataConsumer               As IGenericTickListener
Private mServiceProviderName        As String

Private mStreamId                   As Long

Private mTickRequested              As Boolean
Private mGotFirstTick               As Boolean

Private mContract                   As IContract
Private mContracts                  As IContracts

Private mTargetContractSpec         As IContractSpecifier

Private WithEvents mFetchTicksTC    As TaskController
Attribute mFetchTicksTC.VB_VarHelpID = -1
Private WithEvents mTickStream      As TradingDO27.TickStream
Attribute mTickStream.VB_VarHelpID = -1

Private mTickCounter                As Long

Private mInputTickfileSpecifier     As ITickfileSpecifier

Private mFirstTick                  As GenericTick
Private mFirstTickNotified          As Boolean

Private mUseSynchronousReads        As Boolean

Private WithEvents mFutureWaiter    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mSessionBuilder             As SessionBuilder
Private mSession                    As Session

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@===============================================================================
' IContractFetchListener Interface Members
'@===============================================================================

Private Sub IContractFetchListener_FetchCancelled(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCancelled"
On Error GoTo Err

mFutureWaiter.Cancel
NotifyEvent TickfileEventCodes.TickfileEventDataSourceNotAvailable, "Contract fetch cancelled"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IContractFetchListener_FetchCompleted"
On Error GoTo Err

processContracts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "IContractFetchListener_FetchFailed"
On Error GoTo Err

mFutureWaiter.Cancel
NotifyEvent TickfileEventCodes.TickfileEventDataSourceNotAvailable, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractFetchListener_NotifyContract(ByVal pCookie As Variant, ByVal pContract As ContractUtils27.IContract)
Const ProcName As String = "IContractFetchListener_NotifyContract"
On Error GoTo Err

Static sFirstContractReceived As Boolean

If Not sFirstContractReceived Then
    sFirstContractReceived = True
Else
    mFutureWaiter.Cancel
    NotifyEvent TickfileEventCodes.TickfileEventAmbiguousContractDetails, "Contract not uniquely specified"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' ITickfileReader Interface Members
'@===============================================================================

Private Sub ITickfileReader_CloseInputFile()
Const ProcName As String = "ITickfileReader_CloseInputFile"
On Error GoTo Err

If mParent Is Nothing Then Exit Sub

mParent.TickfileReaderFinished Me
Finish

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ITickfileReader_FireNextTick() As Boolean
Const ProcName As String = "ITickfileReader_FireNextTick"
On Error GoTo Err

If Not mFirstTickNotified Then
    If mGotFirstTick Then
        mFirstTickNotified = True
        notifyTick mFirstTick
        ITickfileReader_FireNextTick = True
    Else
        ITickfileReader_FireNextTick = False
    End If
Else
    ITickfileReader_FireNextTick = getTick
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ITickfileReader_PrepareTickfile( _
                ByVal pTickfileSpec As ITickfileSpecifier) As Boolean
Const ProcName As String = "ITickfileReader_PrepareTickfile"
On Error GoTo Err

Set mInputTickfileSpecifier = pTickfileSpec

' we retrieve our own contract definition, because the one supplied may not be defined
' in this database. Note that we don't use the short name from the source contract as
' this database may use a different convention for short name (eg "FTSE 100 Future Sep 09"
' instead of "ZU9")
With mInputTickfileSpecifier.Contract.Specifier
    Set mTargetContractSpec = CreateContractSpecifier("", .Symbol, .Exchange, .SecType, .CurrencyCode, .Expiry, .Strike, .Right)
End With

checkFetchContract

ITickfileReader_PrepareTickfile = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get ITickfileReader_ReplayProgressPercent() As Long
ITickfileReader_ReplayProgressPercent = 0
End Property

Private Property Get ITickfileReader_ReplayProgressTicks() As Long
ITickfileReader_ReplayProgressTicks = mTickCounter
End Property

Private Function ITickfileReader_Supports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = ((Capabilities And mCapabilities) = Capabilities)
End Function

Private Property Get ITickfileReader_TickfileSizeBytes() As Long
ITickfileReader_TickfileSizeBytes = 0
End Property

'@===============================================================================
' mFetchTicksTC Event Handlers
'@===============================================================================

Private Sub mFetchTicksTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mFetchTicksTC_Completed"
On Error GoTo Err

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    NotifyEvent TickfileEventCodes.TickfileEventDataSourceNotAvailable, ev.ErrorMessage
ElseIf IsEmpty(ev.Result) Then
    NotifyEvent TickfileEventCodes.TickfileEventNoContractDetails, "Contract not known"
Else
    Set mTickStream = ev.Result
    mTickStream.StartStream True
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@===============================================================================
' mFutureWaiter Event Handlers
'@===============================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is TradingDB Then
    Set mTradingDB = ev.Future.Value
    
    checkFetchContract
ElseIf TypeOf ev.Future.Value Is IContracts Then
    Set mContracts = ev.Future.Value
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@===============================================================================
' mTickStream Event Handlers
'@===============================================================================

Private Sub mTickStream_DataAvailable()
Const ProcName As String = "mTickStream_DataAvailable"
On Error GoTo Err

logMessage "Ticks available", ProcName, , LogLevelDetail
If mTickRequested Or Not mGotFirstTick Then
    mTickRequested = False
    getTick
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub Finish()
Const ProcName As String = "finish"
On Error GoTo Err

Set mTickStream = Nothing
Set mReplayMonitor = Nothing
Set mDataConsumer = Nothing
Set mParent = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pStreamId As Long, _
                ByVal pTradingDBFuture As IFuture, _
                ByVal pReplayMonitor As ITickfileReplayMonitor, _
                ByVal pParent As TickfileStore, _
                ByVal pDataConsumer As IGenericTickListener, _
                ByVal pCapabilities As Long, _
                ByVal pUseSynchronousReads As Boolean)
Const ProcName As String = "initialise"
On Error GoTo Err

mStreamId = pStreamId
Set mReplayMonitor = pReplayMonitor
Set mDataConsumer = pDataConsumer
Set mParent = pParent

If pTradingDBFuture.IsAvailable Then
    Set mTradingDB = pTradingDBFuture.Value
Else
    mFutureWaiter.Add pTradingDBFuture
End If

mCapabilities = pCapabilities
mUseSynchronousReads = pUseSynchronousReads

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub checkFetchContract()
Const ProcName As String = "checkFetchContract"
On Error GoTo Err

If mTradingDB Is Nothing Then Exit Sub

If mUseSynchronousReads Then
    logMessage "Fetching contract synchronously", ProcName, mTargetContractSpec.ToString
    Set mContracts = mTradingDB.FetchContracts(mTargetContractSpec)
    processContracts
    If Not mTickStream Is Nothing Then
        mTickStream.StartStream False
        getTick
    End If
Else
    logMessage "Fetching contract asynchronously", ProcName, mTargetContractSpec.ToString
    mFutureWaiter.Add mTradingDB.FetchContractsAsync(mTargetContractSpec, Me)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getTick() As Boolean
Const ProcName As String = "getTick"
On Error GoTo Err

If mTickStream Is Nothing Then
    ' tick reader has been Finished (eg because TickfileStore has been finished)
    Exit Function
End If

Dim tick As GenericTick
Do While mTickStream.GetNextTick(tick)
    If tick.TimeStamp >= mInputTickfileSpecifier.FromDate Then
        If tick.TimeStamp >= mInputTickfileSpecifier.ToDate Then
            noMoreTicks
            Exit Function
        End If
        
        If Not mInputTickfileSpecifier.EntireSession Then
            processTick tick
            getTick = True
            Exit Function
        Else
            mSessionBuilder.SetSessionCurrentTime tick.TimeStamp
            If mSession.IsTimeInSession(tick.TimeStamp) Then
                processTick tick
                getTick = True
                Exit Function
            End If
        End If
    End If
Loop
    
If mTickStream.Eos Then
    noMoreTicks
Else
    logMessage "Waiting for ticks", ProcName, , LogLevelHighDetail
    mTickRequested = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLogger.Log pMsg, pProcName, ModuleName, pLogLevel, pMsgQualifier
End Sub

Private Sub noMoreTicks()
Const ProcName As String = "noMoreTicks"
On Error GoTo Err

If mGotFirstTick Then
    mReplayMonitor.NotifyReplayComplete Me, mTickCounter
Else
    NotifyEvent TickfileEventCodes.TickfileEventFileIsEmpty, "No data for this specifier"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyEvent(ByVal pEventCode As Long, ByVal pEventMessage As String)
Const ProcName As String = "NotifyEvent"
On Error GoTo Err

Dim lEv As NotificationEventData
Set lEv.Source = mParent

lEv.EventCode = pEventCode
lEv.EventMessage = pEventMessage
mReplayMonitor.NotifyEvent lEv

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyTick( _
                ByRef pTick As GenericTick)
Const ProcName As String = "notifyTick"
On Error GoTo Err

Dim ev As GenericTickEventData
Set ev.Source = mParent
ev.StreamId = mStreamId
ev.tick = pTick
mDataConsumer.notifyTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processContracts()
Const ProcName As String = "processContracts"
On Error GoTo Err

Dim ev As NotificationEventData
Set ev.Source = mParent

If mContracts.Count = 0 Then
    NotifyEvent TickfileEventCodes.TickfileEventNoContractDetails, "Contract not known"
ElseIf mContracts.Count > 1 Then
    NotifyEvent TickfileEventCodes.TickfileEventAmbiguousContractDetails, "Contract not unique"
Else
    logMessage "contract fetched", ProcName, , LogLevelHighDetail
    Set mContract = mContracts.ItemAtIndex(1)
    
    With mInputTickfileSpecifier
        .Contract = mContract
        If .EntireSession Then
            Set mSessionBuilder = CreateSessionBuilder(mContract.SessionStartTime, mContract.SessionEndTime, GetTimeZone(mContract.TimeZoneName))
            Set mSession = mSessionBuilder.Session
            
            If .ToDate = 0# Then .ToDate = .FromDate
            
            Dim st As SessionTimes
            st = GetSessionTimes(CDbl(Int(.FromDate)) + mContract.SessionStartTime + 1# / 86400#, _
                                mContract.SessionStartTime, _
                                mContract.SessionEndTime)
            .FromDate = st.StartTime
            st = GetSessionTimes(CDbl(Int(.ToDate)) + mContract.SessionStartTime + 1# / 86400#, _
                                mContract.SessionStartTime, _
                                mContract.SessionEndTime)
            .ToDate = st.EndTime
        End If
    
        If mUseSynchronousReads Then
            logMessage "Fetching ticks synchronously", ProcName, , LogLevelHighDetail
            Set mTickStream = mTradingDB.FetchTicks(mContract.Specifier, .FromDate, .ToDate)
        Else
            logMessage "Fetching ticks asynchronously", ProcName, , LogLevelHighDetail
            Set mFetchTicksTC = mTradingDB.FetchTicksAsync(mContract.Specifier, .FromDate, .ToDate)
        End If
    End With
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTick(ByRef pTick As GenericTick)
Const ProcName As String = "processTick"
On Error GoTo Err

mTickCounter = mTickCounter + 1
If Not mGotFirstTick Then
    mGotFirstTick = True
    mFirstTick = pTick
    logMessage "Notify ready", ProcName, , LogLevelHighDetail
    mReplayMonitor.NotifyReady Me, mInputTickfileSpecifier
Else
    notifyTick pTick
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
