VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IGenericTickSource
Implements IMarketDataSource

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

Private Const ModuleName                            As String = "MarketDataSource"

Private Const ConfigSettingIsMarketDataRequested    As String = "&IsMarketDataRequested"
Private Const ConfigSettingIsMarketDepthRequested   As String = "&IsMarketDepthRequested"

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManagerRef                       As WeakReference
Private mDataSourceBaseRef                          As WeakReference

Private mContractFuture                             As IFuture

Private mStudyBase                                  As StudyBaseForTickDataInput

Private mGenericTickListeners                       As Listeners
Private mCurrentGenericTickListeners()              As Object
Private mSomeCurrentGenericTickListeners            As Boolean

Private mMarketDepthListeners                       As Listeners
Private mCurrentMarketDepthListeners()              As Object
Private mSomeCurrentMarketDepthListeners            As Boolean

Private mOptionComputationListeners                As Listeners
Private mCurrentOptionComputationListeners()       As Object
Private mSomeCurrentOptionComputationListeners     As Boolean

Private mPriceChangeListeners                       As Listeners
Private mCurrentPriceChangeListeners()              As Object
Private mSomeCurrentPriceChangeListeners            As Boolean

Private mQuoteListeners                             As Listeners
Private mCurrentQuoteListeners()                    As Object
Private mSomeCurrentQuoteListeners                  As Boolean

Private mRawMarketDepthListeners                    As Listeners
Private mCurrentRawMarketDepthListeners()           As Object
Private mSomeCurrentRawMarketDepthListeners         As Boolean

Private mStateChangeListeners                       As Listeners
Private mCurrentStateChangeListeners()              As Object
Private mSomeCurrentStateChangeListeners            As Boolean

Private mErrorListeners                             As Listeners
Private mCurrentErrorListeners()                    As Object
Private mSomeCurrentErrorListeners                  As Boolean


Private mTag                                        As Variant

Private mState                                      As MarketDataSourceStates

Private mHandle                                     As Long

Private mKey                                        As String

Private mSessionFutureBuilder                       As New FutureBuilder
Attribute mSessionFutureBuilder.VB_VarHelpID = -1

Private mIsMarketDataRequested                      As Boolean
Private mIsMarketDepthRequested                     As Boolean

Private mSubscribedToMarketData                     As Boolean
Private mSubscribedToMarketDepth                    As Boolean

Private mConfig                                     As ConfigurationSection

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mStreamId                                   As Long

Private mClockFutureBuilder                         As New FutureBuilder

Private mTickfileStreamId                           As Long

Private mErrorMessage                               As String

Private mPriorityTickListener                       As IGenericTickListener

Private mDontSetReady                               As Boolean

Private mIsSnapshot                                 As Boolean
Private mSnapshotRefreshInterval                    As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

mHandle = gAllocateHandle
Set mErrorListeners = New Listeners
Set mGenericTickListeners = New Listeners
Set mMarketDepthListeners = New Listeners
Set mRawMarketDepthListeners = New Listeners
Set mPriceChangeListeners = New Listeners
Set mQuoteListeners = New Listeners
Set mStateChangeListeners = New Listeners

Set mFutureWaiter = New FutureWaiter

mTickfileStreamId = NullIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
On Error Resume Next
gLogger.Log "MarketDataSource terminated (" & mKey & ")", ProcName, ModuleName, LogLevelDetail
End Sub

'@================================================================================
' IGenericTickSource Interface Members
'@================================================================================

Private Sub IGenericTickSource_AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IGenericTickSource_AddGenericTickListener"
On Error GoTo Err

AddGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IGenericTickSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IGenericTickSource_RemoveGenericTickListener"
On Error GoTo Err

RemoveGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDataSource Interface Members
'@================================================================================

Private Sub IMarketDataSource_AddErrorListener(ByVal pListener As IErrorListener)
Const ProcName As String = "IMarketDataSource_AddErrorListener"
On Error GoTo Err

AddErrorListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_AddGenericTickListener"
On Error GoTo Err

AddGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_AddMarketDepthListener(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "IMarketDataSource_AddMarketDepthListener"
On Error GoTo Err

AddMarketDepthListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_AddOptionComputationListener(ByVal pListener As IOptionComputationListener)
Const ProcName As String = "IMarketDataSource_AddOptionComputationListener"
On Error GoTo Err

AddOptionComputationListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_AddPriceChangeListener(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "IMarketDataSource_AddPriceChangeListener"
On Error GoTo Err

AddPriceChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_AddQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "IMarketDataSource_AddQuoteListener"
On Error GoTo Err

AddQuoteListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_AddRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "IMarketDataSource_AddRawMarketDepthListener"
On Error GoTo Err

AddRawMarketDepthListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "IMarketDataSource_AddStateChangeListener"
On Error GoTo Err

AddStateChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IMarketDataSource_ClockFuture() As IFuture
Set IMarketDataSource_ClockFuture = ClockFuture
End Property

Private Property Get IMarketDataSource_ContractFuture() As IFuture
Set IMarketDataSource_ContractFuture = ContractFuture
End Property

Private Property Get IMarketDataSource_CurrentQuote(ByVal pTickType As TickTypes) As Quote
Const ProcName As String = "IMarketDataSource_CurrentQuote"
On Error GoTo Err

IMarketDataSource_CurrentQuote = CurrentQuote(pTickType)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IMarketDataSource_CurrentTick(ByVal pTickType As TickTypes) As GenericTick
IMarketDataSource_CurrentTick = CurrentTick(pTickType)
End Property

Private Property Get IMarketDataSource_ErrorMessage() As String
IMarketDataSource_ErrorMessage = ErrorMessage
End Property

Private Sub IMarketDataSource_Finish()
Const ProcName As String = "IMarketDataSource_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IMarketDataSource_Handle() As Long
Const ProcName As String = "IMarketDataSource_Handle"
On Error GoTo Err

IMarketDataSource_Handle = Handle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IMarketDataSource_HasCurrentTick(ByVal pTickType As TickUtils27.TickTypes) As Boolean
IMarketDataSource_HasCurrentTick = HasCurrentTick(pTickType)
End Property

Private Property Get IMarketDataSource_IsFinished() As Boolean
IMarketDataSource_IsFinished = IsFinished
End Property

Private Property Get IMarketDataSource_IsMarketDataRequested() As Boolean
IMarketDataSource_IsMarketDataRequested = IsMarketDataRequested
End Property

Private Property Get IMarketDataSource_IsMarketDepthRequested() As Boolean
IMarketDataSource_IsMarketDepthRequested = IsMarketDepthRequested
End Property

Private Property Get IMarketDataSource_IsTickReplay() As Boolean
IMarketDataSource_IsTickReplay = IsTickReplay
End Property

Private Property Get IMarketDataSource_Key() As String
IMarketDataSource_Key = Key
End Property

Private Sub IMarketDataSource_NotifyCurrentDOM(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "IMarketDataSource_NotifyCurrentDOM"
On Error GoTo Err

NotifyCurrentDOM pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener)
Const ProcName As String = "IMarketDataSource_NotifyCurrentMarketDataQuoteValues"
On Error GoTo Err

NotifyCurrentMarketDataQuoteValues pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_NotifyCurrentMarketDataTickValues"
On Error GoTo Err

NotifyCurrentMarketDataTickValues pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDepthValues(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "IMarketDataSource_NotifyCurrentMarketDepthValues"
On Error GoTo Err

NotifyCurrentMarketDepthValues pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IMarketDataSource_NumberOfMarketDepthRows() As Long
IMarketDataSource_NumberOfMarketDepthRows = NumberOfMarketDepthRows
End Property

Private Property Get IMarketDataSource_Options() As MarketDataSourceOptions
IMarketDataSource_Options = Options
End Property

Private Property Get IMarketDataSource_PriceChange() As PriceChange
IMarketDataSource_PriceChange = PriceChange
End Property

Private Property Let IMarketDataSource_PriorityTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_PriorityTickListener"
On Error GoTo Err

PriorityTickListener = pListener

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IMarketDataSource_Refresh()
Const ProcName As String = "IMarketDataSource_Refresh"
On Error GoTo Err

Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RefreshPriceChange(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "IMarketDataSource_RefreshPriceChange"
On Error GoTo Err

RefreshPriceChange pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RefreshQuotes(ByVal pListener As IQuoteListener)
Const ProcName As String = "IMarketDataSource_RefreshQuotes"
On Error GoTo Err

RefreshQuotes pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RefreshTicks(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_RefreshTicks"
On Error GoTo Err

RefreshTicks pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RemoveErrorListener(ByVal pListener As IErrorListener)
Const ProcName As String = "IMarketDataSource_RemoveErrorListener"
On Error GoTo Err

RemoveErrorListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RemoveFromConfig()
Const ProcName As String = "IMarketDataSource_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_RemoveGenericTickListener"
On Error GoTo Err

RemoveGenericTickListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RemoveMarketDepthListener(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "IMarketDataSource_RemoveMarketDepthListener"
On Error GoTo Err

RemoveMarketDepthListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RemoveOptionComputationListener(ByVal pListener As IOptionComputationListener)
Const ProcName As String = "IMarketDataSource_RemoveOptionComputationListener"
On Error GoTo Err

RemoveOptionComputationListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RemovePriceChangeListener(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "IMarketDataSource_RemovePriceChangeListener"
On Error GoTo Err

RemovePriceChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RemoveQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "IMarketDataSource_RemoveQuoteListener"
On Error GoTo Err

RemoveQuoteListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataSource_RemoveRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
RemoveRawMarketDepthListener pListener
End Sub

Private Sub IMarketDataSource_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "IMarketDataSource_RemoveStateChangeListener"
On Error GoTo Err

RemoveStateChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IMarketDataSource_SessionFuture() As IFuture
Set IMarketDataSource_SessionFuture = SessionFuture
End Property

Private Sub IMarketDataSource_StartMarketData()
StartMarketData
End Sub

Private Sub IMarketDataSource_StartMarketDepth()
StartMarketDepth
End Sub

Private Property Get IMarketDataSource_State() As MarketDataSourceStates
Const ProcName As String = "IMarketDataSource_State"
On Error GoTo Err

IMarketDataSource_State = State

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IMarketDataSource_StopMarketData()
StopMarketData
End Sub

Private Sub IMarketDataSource_StopMarketDepth()
StopMarketDepth
End Sub

Public Property Get IMarketDataSource_StudyBase() As IStudyBase
Set IMarketDataSource_StudyBase = StudyBase
End Property

Private Property Get IMarketDataSource_Timestamp() As Date
Const ProcName As String = "IMarketDataSource_Timestamp"
On Error GoTo Err

IMarketDataSource_Timestamp = Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

If Not mDontSetReady Then setReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsPending Then Exit Sub

Static sGotContract As Boolean
Static sGotClock As Boolean

If ev.Future.IsFaulted Then
    SetError ev.Future.ErrorNumber, ev.ContinuationData & " creation failed: " & ev.Future.ErrorMessage
    mFutureWaiter.Cancel
ElseIf ev.Future.IsCancelled Then
    SetError ErrorCodes.ErrIllegalStateException, ev.ContinuationData & " creation cancelled"
    mFutureWaiter.Cancel
ElseIf TypeOf ev.Future.Value Is IContract Then
    gLogger.Log "Contract available for MarketDataSource with key: " & mKey, ProcName, ModuleName, LogLevelHighDetail
    sGotContract = True
    If sGotClock Then
        If Not processContractAvailable Then mDontSetReady = True
    End If
ElseIf TypeOf ev.Future.Value Is Clock Then
    gLogger.Log "Clock available for MarketDataSource with key: " & mKey, ProcName, ModuleName, LogLevelHighDetail
    sGotClock = True
    mClockFutureBuilder.Value = ev.Future.Value
    mClockFutureBuilder.Complete
    If sGotContract Then
        If Not processContractAvailable Then mDontSetReady = True
    End If
ElseIf TypeOf ev.Future.Value Is StudyManager Then
    gLogger.Log "StudyManager available for MarketDataSource with key: " & Key, ProcName, ModuleName, LogLevelHighDetail
    ' note that this will always be after the other futures have become
    ' available because it depends on the session future being available,
    ' which only occurs during processContractAvailable()
    Set mStudyBase = New StudyBaseForTickDataInput
    mStudyBase.Initialise ev.Future.Value, mContractFuture
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ClockFuture() As IFuture
Set ClockFuture = mClockFutureBuilder.Future
End Property

Public Property Get ContractFuture() As IFuture
Set ContractFuture = mContractFuture
End Property

Public Property Get CurrentQuote(ByVal pTickType As TickTypes) As Quote
Const ProcName As String = "CurrentQuote"
On Error GoTo Err

checkNotFinished
Assert mState = MarketDataSourceStateRunning, "Not in Running state"
CurrentQuote = DataSourceBase.GetCurrentQuote(pTickType)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentTick(ByVal pTickType As TickTypes) As GenericTick
Const ProcName As String = "CurrentTick"
On Error GoTo Err

checkNotFinished

Assert mState = MarketDataSourceStateRunning, "Not in Running state"
CurrentTick = DataSourceBase.GetCurrentTick(pTickType)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ErrorMessage() As String
ErrorMessage = mErrorMessage
End Property

Public Property Get Handle() As Long
' we don't call checkNotFinished because the application may handle
' the StateChange event when Finish is called and need the handle then
Handle = mHandle
End Property

Public Property Get HasCurrentTick(ByVal pTickType As TickTypes) As Boolean
Const ProcName As String = "HasCurrentTick"
On Error GoTo Err

checkNotFinished

Assert mState = MarketDataSourceStateRunning, "Not in Running state"
HasCurrentTick = DataSourceBase.HasCurrentTick(pTickType)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get HasQuoteListeners()
HasQuoteListeners = (mQuoteListeners.Count <> 0)
End Property

Friend Property Get HasPriceChangeListeners()
HasPriceChangeListeners = (mPriceChangeListeners.Count <> 0)
End Property

Public Property Get IsFinished() As Boolean
IsFinished = (mState = MarketDataSourceStateFinished)
End Property

Public Property Get IsMarketDataRequested() As Boolean
IsMarketDataRequested = mIsMarketDataRequested
End Property

Public Property Get IsMarketDepthRequested() As Boolean
IsMarketDepthRequested = mIsMarketDepthRequested
End Property

Public Property Get IsTickReplay() As Boolean
IsTickReplay = (mTickfileStreamId <> NullIndex)
End Property

Public Property Get Key() As String
Key = mKey
End Property

Private Property Get NumberOfMarketDepthRows() As Long
NumberOfMarketDepthRows = NumberOfMarketDepthRows
End Property

Private Property Get Options() As MarketDataSourceOptions
Options = Options
End Property

Public Property Get PriceChange() As PriceChange
Const ProcName As String = "PriceChange"
On Error GoTo Err

checkNotFinished
PriceChange = DataSourceBase.GetPriceChange()

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PriorityTickListener(ByVal Value As IGenericTickListener)
Const ProcName As String = "PriorityTickListener"
On Error GoTo Err

Set mPriorityTickListener = Value
If Not mDataSourceBaseRef Is Nothing Then DataSourceBase.AddPriorityTickListener mPriorityTickListener

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SessionFuture() As IFuture
Set SessionFuture = mSessionFutureBuilder.Future
End Property

Public Property Get State() As MarketDataSourceStates
State = mState
End Property

Friend Property Get StreamId() As Long
Const ProcName As String = "StreamId"
On Error GoTo Err

StreamId = mStreamId

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyBase() As IStudyBase
Set StudyBase = mStudyBase
End Property

Friend Property Get SubscribedToMarketData() As Boolean
SubscribedToMarketData = mSubscribedToMarketData
End Property

Friend Property Get SubscribedToMarketDepth() As Boolean
SubscribedToMarketDepth = mSubscribedToMarketDepth
End Property

Public Property Let Tag(ByVal Value As Variant)
Const ProcName As String = "Tag"
On Error GoTo Err

checkNotFinished
gSetVariant mTag, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Tag() As Variant
Const ProcName As String = "Tag"
On Error GoTo Err

gSetVariant Tag, mTag

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Timestamp() As Date
Const ProcName As String = "Timestamp"
On Error GoTo Err

checkNotFinished
Assert mState = MarketDataSourceStateReady Or mState = MarketDataSourceStateRunning, "Not in Ready or Running state"
Timestamp = DataSourceBase.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddErrorListener(ByVal pListener As IErrorListener)
Const ProcName As String = "AddErrorListener"
On Error GoTo Err

checkNotFinished
mErrorListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddGenericTickListener"
On Error GoTo Err

checkNotFinished
mGenericTickListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddMarketDepthListener(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "AddMarketDepthListener"
On Error GoTo Err

checkNotFinished
mMarketDepthListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddOptionComputationListener(ByVal pListener As IOptionComputationListener)
Const ProcName As String = "Add"
On Error GoTo Err

mOptionComputationListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddPriceChangeListener(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "AddPriceChangeListener"
On Error GoTo Err

checkNotFinished
mPriceChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "AddQuoteListener"
On Error GoTo Err

checkNotFinished
mQuoteListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "AddRawMarketDepthListener"
On Error GoTo Err

checkNotFinished
mRawMarketDepthListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

checkNotFinished
mStateChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

checkNotFinished

If State = MarketDataSourceStateError Then Exit Sub

DataSourceBase.RemoveMarketDataSource Me
setState MarketDataSourceStateFinished

mGenericTickListeners.Clear
mMarketDepthListeners.Clear
mPriceChangeListeners.Clear
mQuoteListeners.Clear
mRawMarketDepthListeners.Clear
mStateChangeListeners.Clear
mErrorListeners.Clear

Set mStudyBase = Nothing

Dim s As String
If mContractFuture.IsAvailable Then
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    s = lContract.Specifier.ToString
End If

gReleaseHandle mHandle
gLogger.Log "MarketDataSource finished (" & mKey & ")", ProcName, ModuleName, , s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireError(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "FireError"
On Error GoTo Err

Dim ev As ErrorEventData
Set ev.Source = Me
ev.ErrorCode = pErrorCode
ev.ErrorMessage = pErrorMessage

updateCurrentErrorListeners

If mSomeCurrentErrorListeners Then
    Dim lListener As IErrorListener
    Dim i As Long
    For i = 0 To UBound(mCurrentErrorListeners)
        Set lListener = mCurrentErrorListeners(i)
        lListener.Notify ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireNoMoreTicks()
Const ProcName As String = "FireNoMoreTicks"
On Error GoTo Err

If mSubscribedToMarketDepth Or mSubscribedToMarketData Then Exit Sub

Dim ev As GenericTickEventData
Set ev.Source = Me
ev.StreamId = mStreamId

UpdateCurrentGenericTickListeners

If mSomeCurrentGenericTickListeners Then
    Dim lListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(mCurrentGenericTickListeners)
        Set lListener = mCurrentGenericTickListeners(i)
        lListener.NoMoreTicks ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FirePriceChange(ByRef pPriceChange As PriceChange)
Const ProcName As String = "FirePriceChange"
On Error GoTo Err

If Not mSubscribedToMarketData Then Exit Sub

Dim ev As PriceChangeEventData
Set ev.Source = Me
ev.PriceChange = pPriceChange

updateCurrentPriceChangeListeners

If mSomeCurrentPriceChangeListeners Then
    Dim lListener As IPriceChangeListener
    Dim i As Long
    For i = 0 To UBound(mCurrentPriceChangeListeners)
        Set lListener = mCurrentPriceChangeListeners(i)
        lListener.Change ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireQuote(ByRef pQuote As Quote, ByVal pTickType As TickTypes)
Const ProcName As String = "FireQuote"
On Error GoTo Err

If Not mSubscribedToMarketData Then Exit Sub

Dim ev As QuoteEventData
Set ev.Source = Me
ev.Quote = pQuote

updateCurrentQuoteListeners

If mSomeCurrentQuoteListeners Then
    Dim lListener As IQuoteListener
    Dim i As Long
    For i = 0 To UBound(mCurrentQuoteListeners)
        Set lListener = mCurrentQuoteListeners(i)
                Select Case pTickType
                Case TickTypeBid
                        lListener.Bid ev
                Case TickTypeAsk
                        lListener.Ask ev
                Case TickTypeClosePrice
                        lListener.PreviousClose ev
                Case TickTypeHighPrice
                        lListener.High ev
                Case TickTypeLowPrice
                        lListener.Low ev
                Case TickTypeMarketDepth
                Case TickTypeMarketDepthReset
                Case TickTypeTrade
                        lListener.Trade ev
                Case TickTypeVolume
                        lListener.Volume ev
                Case TickTypeOpenInterest
                        lListener.OpenInterest ev
                Case TickTypeOpenPrice
                        lListener.SessionOpen ev
                End Select
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireRawMarketDepth( _
                ByRef pTick As GenericTick)
Const ProcName As String = "FireRawMarketDepth"
On Error GoTo Err

If Not mSubscribedToMarketDepth Then Exit Sub

Dim ev As RawMarketDepthEventData
Set ev.Source = Me

If pTick.TickType = TickTypeMarketDepth Then
    ev.MarketMaker = pTick.MarketMaker
    ev.Operation = pTick.Operation
    ev.Position = pTick.Position
    ev.Price = pTick.Price
    ev.Side = pTick.Side
    ev.Size = pTick.Size
End If

updateRawMarketDepthListeners

If mSomeCurrentRawMarketDepthListeners Then
    Dim lListener As IRawMarketDepthListener
    Dim i As Long
    For i = 0 To UBound(mCurrentRawMarketDepthListeners)
        Set lListener = mCurrentRawMarketDepthListeners(i)
        If pTick.TickType = TickTypeMarketDepth Then
                lListener.UpdateMarketDepth ev
        Else
                lListener.ResetMarketDepth ev
        End If
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireResetMarketDepth()
Const ProcName As String = "FireResetMarketDepth"
On Error GoTo Err

If Not mSubscribedToMarketDepth Then Exit Sub

Dim evRaw As RawMarketDepthEventData
Set evRaw.Source = Me

updateRawMarketDepthListeners

If mSomeCurrentRawMarketDepthListeners Then
    Dim lRawListener As IRawMarketDepthListener
    Dim i As Long
    For i = 0 To UBound(mCurrentRawMarketDepthListeners)
        Set lRawListener = mCurrentRawMarketDepthListeners(i)
        lRawListener.ResetMarketDepth evRaw
    Next
End If

Dim ev As MarketDepthEventData
Set ev.Source = Me

updateCurrentMarketDepthListeners

If mSomeCurrentMarketDepthListeners Then
    Dim lListener As IMarketDepthListener
    Dim j As Long
    For j = 0 To UBound(mCurrentMarketDepthListeners)
        Set lListener = mCurrentMarketDepthListeners(j)
        lListener.ResetMarketDepth ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireSetMarketDepthCell( _
                ByVal pSide As DOMSides, _
                ByVal pPrice As Double, _
                ByVal pSize As Long)
Const ProcName As String = "FireSetMarketDepthCell"
On Error GoTo Err

If Not mSubscribedToMarketDepth Then Exit Sub

Dim ev As MarketDepthEventData
Set ev.Source = Me
ev.Price = pPrice
ev.Type = pSide
ev.Size = pSize

updateCurrentMarketDepthListeners
If mSomeCurrentMarketDepthListeners Then
    Dim lListener As IMarketDepthListener
    Dim i As Long
    For i = 0 To UBound(mCurrentMarketDepthListeners)
        Set lListener = mCurrentMarketDepthListeners(i)
        lListener.SetMarketDepthCell ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireOptionComputation(ByRef pOptionComputation As OptionComputation)
Const ProcName As String = "FireOptionComputation"
On Error GoTo Err

If Not mSubscribedToMarketData Then Exit Sub

Dim ev As OptionComputationEventData
Set ev.Source = Me
ev.StreamId = mStreamId
ev.Computation = pOptionComputation

updateCurrentOptionComputationListeners
If mSomeCurrentOptionComputationListeners Then
    Dim lListener As IOptionComputationListener
    Dim i As Long
    For i = 0 To UBound(mCurrentOptionComputationListeners)
        Set lListener = mCurrentOptionComputationListeners(i)
        lListener.NotifyComputation ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireTick(ByRef pTick As GenericTick)
Const ProcName As String = "FireTick"
On Error GoTo Err

If pTick.TickType = TickTypeMarketDepth Or pTick.TickType = TickTypeMarketDepthReset Then
    If Not mSubscribedToMarketDepth Then Exit Sub
Else
    If Not mSubscribedToMarketData Then Exit Sub
End If

' studies must be updated before notifying other listeners
If Not mStudyBase Is Nothing Then
    mStudyBase.NotifyTick pTick
End If

Dim ev As GenericTickEventData
Set ev.Source = Me
ev.StreamId = mStreamId
ev.Tick = pTick

UpdateCurrentGenericTickListeners
If mSomeCurrentGenericTickListeners Then
    Dim lListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(mCurrentGenericTickListeners)
        Set lListener = mCurrentGenericTickListeners(i)
        lListener.NotifyTick ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pMarketDataManager As MarketDataManager, _
                ByVal pKey As String, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pStudyManagerFuture As IFuture, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pTickfileStreamId As Long, _
                ByVal pIsSnapshot As Boolean, _
                ByVal pSnapshotRefreshInterval As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

gLogger.Log "Initialising MarketDataSource with key: " & pKey, ProcName, ModuleName, LogLevelHighDetail

Set mMarketDataManagerRef = CreateWeakReference(pMarketDataManager)

mKey = pKey
Set mContractFuture = pContractFuture
Set mConfig = pConfig
mTickfileStreamId = pTickfileStreamId
mIsSnapshot = pIsSnapshot
mSnapshotRefreshInterval = pSnapshotRefreshInterval

setState MarketDataSourceStateCreated

If mContractFuture.IsAvailable And pClockFuture.IsAvailable Then
    gLogger.Log "Contract available for MarketDataSource with key: " & mKey, ProcName, ModuleName, LogLevelHighDetail
    gLogger.Log "Clock available for MarketDataSource with key: " & mKey, ProcName, ModuleName, LogLevelHighDetail
    mClockFutureBuilder.Value = pClockFuture.Value
    mClockFutureBuilder.Complete
    If Not processContractAvailable Then Exit Sub
    If pStudyManagerFuture Is Nothing Then
        setReady
    Else
        mFutureWaiter.Add pStudyManagerFuture
    End If
Else
    mFutureWaiter.Add pClockFuture
    mFutureWaiter.Add mContractFuture
    If Not pStudyManagerFuture Is Nothing Then mFutureWaiter.Add pStudyManagerFuture
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentDOM( _
                ByVal pListener As IMarketDepthListener)
Const ProcName As String = "NotifyCurrentDOM"
On Error GoTo Err

checkNotFinished
If mState = MarketDataSourceStateRunning Then DataSourceBase.DOMManager.NotifyCurrentDOM pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener)
Const ProcName As String = "NotifyCurrentMarketDataQuoteValues"
On Error GoTo Err

checkNotFinished
If mState = MarketDataSourceStateRunning Then DataSourceBase.NotifyCurrentMarketDataQuoteValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener)
Const ProcName As String = "NotifyCurrentMarketDataTickValues"
On Error GoTo Err

checkNotFinished
If mState = MarketDataSourceStateRunning Then DataSourceBase.NotifyCurrentMarketDataTickValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDepthValues( _
                ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "NotifyCurrentMarketDepthValues"
On Error GoTo Err

checkNotFinished
If mState = MarketDataSourceStateRunning Then DataSourceBase.DOMManager.NotifyCurrentMarketDepthValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Refresh()
Const ProcName As String = "Refresh"
On Error GoTo Err

checkNotFinished

If Not (State = MarketDataSourceStateReady Or State = MarketDataSourceStateRunning) Then Exit Sub

UpdateCurrentGenericTickListeners

If mSomeCurrentGenericTickListeners Then
    Dim lTickListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(mCurrentGenericTickListeners)
        Set lTickListener = mCurrentGenericTickListeners(i)
        DataSourceBase.NotifyCurrentMarketDataTickValues lTickListener, Me
    Next
End If

updateCurrentQuoteListeners

If mSomeCurrentQuoteListeners Then
    Dim lQuoteListener As IQuoteListener
    Dim j As Long
    For j = 0 To UBound(mCurrentQuoteListeners)
        Set lQuoteListener = mCurrentQuoteListeners(j)
        DataSourceBase.NotifyCurrentMarketDataQuoteValues lQuoteListener, Me
    Next
End If

Dim ev As PriceChangeEventData
Set ev.Source = Me
ev.PriceChange = DataSourceBase.GetPriceChange

updateCurrentPriceChangeListeners

If mSomeCurrentPriceChangeListeners Then
    Dim lPricechangeListener As IPriceChangeListener
    Dim k As Long
    For k = 0 To UBound(mCurrentPriceChangeListeners)
        Set lPricechangeListener = mCurrentPriceChangeListeners(k)
        lPricechangeListener.Change ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RefreshPriceChange(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "RefreshPriceChange"
On Error GoTo Err

checkNotFinished
Dim lPricechangeListener As IPriceChangeListener
Dim ev As PriceChangeEventData
Set ev.Source = Me
ev.PriceChange = DataSourceBase.GetPriceChange
pListener.Change ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RefreshQuotes(ByVal pListener As IQuoteListener)
Const ProcName As String = "RefreshQuotes"
On Error GoTo Err

checkNotFinished
DataSourceBase.NotifyCurrentMarketDataQuoteValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RefreshTicks(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RefreshTicks"
On Error GoTo Err

checkNotFinished
DataSourceBase.NotifyCurrentMarketDataTickValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveErrorListener(ByVal pListener As IErrorListener)
Const ProcName As String = "RemoveErrorListener"
On Error GoTo Err

If IsFinished Then Exit Sub

mErrorListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

mConfig.Remove
Set mConfig = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveGenericTickListener"
On Error GoTo Err

If IsFinished Then Exit Sub

mGenericTickListeners.Remove pListener

unsubscribeIfNoMarketDataListeners

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveMarketDepthListener(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "RemoveMarketDepthListener"
On Error GoTo Err

If IsFinished Then Exit Sub

mMarketDepthListeners.Remove pListener

unsubscribeIfNoMarketDepthListeners

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveOptionComputationListener(ByVal pListener As IOptionComputationListener)
Const ProcName As String = "RemoveOptionComputationListener"
On Error GoTo Err

mOptionComputationListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemovePriceChangeListener(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "RemovePriceChangeListener"
On Error GoTo Err

If IsFinished Then Exit Sub

mPriceChangeListeners.Remove pListener

unsubscribeIfNoMarketDataListeners

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "RemoveQuoteListener"
On Error GoTo Err

If IsFinished Then Exit Sub

mQuoteListeners.Remove pListener

unsubscribeIfNoMarketDataListeners

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "RemoveRawMarketDepthListener"
On Error GoTo Err

If IsFinished Then Exit Sub

mRawMarketDepthListeners.Remove pListener

unsubscribeIfNoMarketDepthListeners

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

If IsFinished Then Exit Sub

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMessage As String)
Const ProcName As String = "SetError"
On Error GoTo Err

FireError pErrorCode, pErrorMessage

mErrorMessage = pErrorMessage
setState MarketDataSourceStateError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartMarketData()
Const ProcName As String = "StartMarketData"
On Error GoTo Err

checkNotFinished
Assert Not mIsMarketDataRequested, "Market data already started"
mIsMarketDataRequested = True
If Not (State = MarketDataSourceStateReady Or State = MarketDataSourceStateRunning) Then Exit Sub
subscribeToMarketData
storeSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartMarketDepth()
Const ProcName As String = "StartMarketDepth"
On Error GoTo Err

checkNotFinished
Assert Not mIsMarketDepthRequested, "Market depth already started"
mIsMarketDepthRequested = True
If Not (State = MarketDataSourceStateReady Or State = MarketDataSourceStateRunning) Then Exit Sub
subscribeToMarketDepth
storeSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopMarketData()
Const ProcName As String = "StopMarketData"
On Error GoTo Err

checkNotFinished
If Not mIsMarketDataRequested Then Exit Sub
If Not mSubscribedToMarketData Then Exit Sub

DataSourceBase.RemoveMarketDataSubscription
mSubscribedToMarketData = False
mIsMarketDataRequested = False

If Not mIsMarketDepthRequested Then setState MarketDataSourceStateStopped

storeSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopMarketDepth()
Const ProcName As String = "StopMarketDepth"
On Error GoTo Err

checkNotFinished
If Not mIsMarketDepthRequested Then Exit Sub
If Not mSubscribedToMarketDepth Then Exit Sub

DataSourceBase.RemoveMarketDepthSubscription
mSubscribedToMarketDepth = False
mIsMarketDepthRequested = False

If Not mIsMarketDataRequested Then setState MarketDataSourceStateStopped

storeSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkNotFinished()
Assert Not IsFinished, "Object has been Finished"
End Sub

Private Function DataSourceBase() As DataSourceBase
Set DataSourceBase = mDataSourceBaseRef.Target
End Function

Private Sub fireStateChange(ByVal pState As MarketDataSourceStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = pState

updateCurrentStateChangeListeners

If mSomeCurrentStateChangeListeners Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(mCurrentStateChangeListeners)
        Set lListener = mCurrentStateChangeListeners(i)
        lListener.Change ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getMarketDataManager() As MarketDataManager
Set getMarketDataManager = mMarketDataManagerRef.Target
End Function

Private Function processContractAvailable() As Boolean
Const ProcName As String = "processContractAvailable"
On Error GoTo Err

Dim lContract As IContract
Set lContract = mContractFuture.Value

Set mDataSourceBaseRef = CreateWeakReference( _
                            getMarketDataManager.GetDataSourceBase( _
                                    mContractFuture, _
                                    mClockFutureBuilder.Future, _
                                    mIsSnapshot, _
                                    mSnapshotRefreshInterval, _
                                    mTickfileStreamId))
If Not mPriorityTickListener Is Nothing Then DataSourceBase.AddPriorityTickListener mPriorityTickListener
DataSourceBase.AddMarketDataSource Me

mStreamId = DataSourceBase.StreamId

mSessionFutureBuilder.Value = DataSourceBase.Session
mSessionFutureBuilder.Complete

If IsContractExpired(lContract) And Not IsTickReplay Then
    SetError ErrorCodes.ErrIllegalStateException, "Contract has expired"
    processContractAvailable = False
Else
    processContractAvailable = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setReady()
Const ProcName As String = "setReady"
On Error GoTo Err

If Not ContractFuture.IsAvailable Then Exit Sub

If State <> MarketDataSourceStateCreated Then Exit Sub

gLogger.Log "MarketDataSource is ready: key: " & mKey, ProcName, ModuleName, LogLevelHighDetail
setState MarketDataSourceStateReady

If mIsMarketDataRequested Then subscribeToMarketData
If mIsMarketDepthRequested Then subscribeToMarketDepth

storeSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setState(ByVal pState As MarketDataSourceStates)
Const ProcName As String = "setState"
On Error GoTo Err

If pState = mState Then Exit Sub

mState = pState
fireStateChange mState

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeSettings()
Const ProcName As String = "storeSettings"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

' don't save settings until we have the contract. Missing contract
' data could cause problems on recovery
If State = MarketDataSourceStateCreated Then Exit Sub

SaveContractToConfig mContractFuture.Value, mConfig.AddConfigurationSection(ConfigSectionContract)
mConfig.SetSetting ConfigSettingIsMarketDataRequested, mIsMarketDataRequested
mConfig.SetSetting ConfigSettingIsMarketDepthRequested, mIsMarketDepthRequested

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub subscribeToMarketData()
Const ProcName As String = "subscribeToMarketData"
On Error GoTo Err

If mSubscribedToMarketData Then Exit Sub
mSubscribedToMarketData = True
setState MarketDataSourceStateRunning
DataSourceBase.AddMarketDataSubscription

Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub subscribeToMarketDepth()
Const ProcName As String = "subscribeToMarketDepth"
On Error GoTo Err

If mSubscribedToMarketDepth Then Exit Sub
mSubscribedToMarketDepth = True
setState MarketDataSourceStateRunning
DataSourceBase.AddMarketDepthSubscription

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unsubscribeIfNoMarketDataListeners()
Const ProcName As String = "unsubscribeIfNoMarketDataListeners"
On Error GoTo Err

If mState = MarketDataSourceStateRunning Then
    If mGenericTickListeners.Count = 0 And mQuoteListeners.Count = 0 And mPriceChangeListeners.Count = 0 Then StopMarketData
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unsubscribeIfNoMarketDepthListeners()
Const ProcName As String = "unsubscribeIfNoMarketDepthListeners"
On Error GoTo Err

If mState = MarketDataSourceStateRunning Then
    If mMarketDepthListeners.Count = 0 And mRawMarketDepthListeners.Count = 0 Then StopMarketDepth
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateCurrentErrorListeners()
Static sInit As Boolean
If Not sInit Or Not mErrorListeners.Valid Then
    sInit = True
    mSomeCurrentErrorListeners = mErrorListeners.GetCurrentListeners(mCurrentErrorListeners)
End If
End Sub

Private Sub UpdateCurrentGenericTickListeners()
Static sInit As Boolean
If Not sInit Or Not mGenericTickListeners.Valid Then
    sInit = True
    mSomeCurrentGenericTickListeners = mGenericTickListeners.GetCurrentListeners(mCurrentGenericTickListeners)
End If
End Sub

Private Sub updateCurrentMarketDepthListeners()
Static sInit As Boolean
If Not sInit Or Not mMarketDepthListeners.Valid Then
    sInit = True
    mSomeCurrentMarketDepthListeners = mMarketDepthListeners.GetCurrentListeners(mCurrentMarketDepthListeners)
End If
End Sub

Private Sub updateCurrentPriceChangeListeners()
Static sInit As Boolean
If Not sInit Or Not mPriceChangeListeners.Valid Then
    sInit = True
    mSomeCurrentPriceChangeListeners = mPriceChangeListeners.GetCurrentListeners(mCurrentPriceChangeListeners)
End If
End Sub

Private Sub updateCurrentOptionComputationListeners()
Static sInit As Boolean
If Not sInit Or Not mOptionComputationListeners.Valid Then
    sInit = True
    mSomeCurrentOptionComputationListeners = mOptionComputationListeners.GetCurrentListeners(mCurrentOptionComputationListeners)
End If
End Sub

Private Sub updateCurrentQuoteListeners()
Static sInit As Boolean
If Not sInit Or Not mQuoteListeners.Valid Then
    sInit = True
    mSomeCurrentQuoteListeners = mQuoteListeners.GetCurrentListeners(mCurrentQuoteListeners)
End If
End Sub

Private Sub updateRawMarketDepthListeners()
Static sInit As Boolean
If Not sInit Or Not mRawMarketDepthListeners.Valid Then
    sInit = True
    mSomeCurrentRawMarketDepthListeners = mRawMarketDepthListeners.GetCurrentListeners(mCurrentRawMarketDepthListeners)
End If
End Sub

Private Sub updateCurrentStateChangeListeners()
Static sInit As Boolean
If Not sInit Or Not mStateChangeListeners.Valid Then
    sInit = True
    mSomeCurrentStateChangeListeners = mStateChangeListeners.GetCurrentListeners(mCurrentStateChangeListeners)
End If
End Sub

