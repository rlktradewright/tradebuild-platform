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

Private mDataStreamRef                              As WeakReference

Private mContract                                   As IContract
Private mContractSpecifier                          As IContractSpecifier

Private mGenericTickListeners                       As EnumerableCollection
Private mMarketDepthListeners                       As EnumerableCollection
Private mPriceChangeListeners                       As EnumerableCollection
Private mQuoteListeners                             As EnumerableCollection
Private mRawMarketDepthListeners                    As EnumerableCollection
Private mStateChangeListeners                       As EnumerableCollection
Private mErrorListeners                             As EnumerableCollection

Private mTag                                        As Variant

Private mState                                      As MarketDataSourceStates

Private mHandle                                     As Long

Private mKey                                        As String

Private mSession                                    As Session
Attribute mSession.VB_VarHelpID = -1

Private mIsMarketDataRequested                      As Boolean
Private mIsMarketDepthRequested                     As Boolean

Private mClock                                      As Clock

Private mConfig                                     As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

mHandle = gAllocateHandle
Set mErrorListeners = New EnumerableCollection
Set mGenericTickListeners = New EnumerableCollection
Set mMarketDepthListeners = New EnumerableCollection
Set mRawMarketDepthListeners = New EnumerableCollection
Set mPriceChangeListeners = New EnumerableCollection
Set mQuoteListeners = New EnumerableCollection
Set mStateChangeListeners = New EnumerableCollection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
gReleaseHandle mHandle
End Sub

'@================================================================================
' IMarketDataSource Interface Members
'@================================================================================

Private Sub IMarketDataSource_AddErrorListener(ByVal pListener As TWUtilities40.ErrorListener)
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

Private Sub IMarketDataSource_AddMarketDepthListener(ByVal pListener As TickUtils27.IMarketDepthListener)
Const ProcName As String = "IMarketDataSource_AddMarketDepthListener"
On Error GoTo Err

AddMarketDepthListener pListener

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

Private Sub IMarketDataSource_AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IMarketDataSource_AddStateChangeListener"
On Error GoTo Err

AddStateChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IMarketDataSource_Clock() As Clock
Set IMarketDataSource_Clock = Clock
End Property

Private Property Get IMarketDataSource_Contract() As IContract
Set IMarketDataSource_Contract = mContract
End Property

Private Property Get IMarketDataSource_ContractSpecifier() As IContractSpecifier
Set IMarketDataSource_ContractSpecifier = ContractSpecifier
End Property

Private Function IMarketDataSource_ConvertTickToQuote(pTick As GenericTick) As TickUtils27.Quote
Const ProcName As String = "IMarketDataSource_ConvertTickToQuote"
On Error GoTo Err

IMarketDataSource_ConvertTickToQuote = ConvertTickToQuote(pTick)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

Private Sub IMarketDataSource_RemoveErrorListener(ByVal pListener As TWUtilities40.ErrorListener)
Const ProcName As String = "IMarketDataSource_RemoveErrorListener"
On Error GoTo Err

RemoveErrorListener pListener

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

Private Sub IMarketDataSource_RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "IMarketDataSource_RemoveStateChangeListener"
On Error GoTo Err

RemoveStateChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IMarketDataSource_Session() As Session
Set IMarketDataSource_Session = Session
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

Private Property Get IMarketDataSource_Timestamp() As Date
Const ProcName As String = "IMarketDataSource_Timestamp"
On Error GoTo Err

IMarketDataSource_Timestamp = Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Clock() As Clock
Set Clock = mClock
End Property

Friend Property Let Contract(ByVal Value As IContract)
Const ProcName As String = "Contract"
On Error GoTo Err

Set mContract = Value
If Not mContract Is Nothing Then
    Set mContractSpecifier = mContract.Specifier
    Set mSession = getDataStream.Session
    Set mClock = getDataStream.Clock
    
    setState MarketDataSourceStateReady
    
    If mIsMarketDataRequested Or mIsMarketDepthRequested Then setState MarketDataSourceStateRunning
    If mIsMarketDataRequested Then Refresh
End If

storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Contract() As IContract
Set Contract = mContract
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Public Property Get CurrentTick(ByVal pTickType As TickTypes) As GenericTick
Const ProcName As String = "CurrentTick"
On Error GoTo Err

checkNotFinished

Assert mState = MarketDataSourceStateRunning, "Not in Running state"
CurrentTick = getDataStream.GetCurrentTick(pTickType)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentQuote(ByVal pTickType As TickTypes) As Quote
Const ProcName As String = "CurrentQuote"
On Error GoTo Err

checkNotFinished
Assert mState = MarketDataSourceStateRunning, "Not in Running state"
CurrentQuote = getDataStream.GetCurrentQuote(pTickType)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ErrorListeners() As EnumerableCollection
Set ErrorListeners = mErrorListeners
End Property

Friend Property Get GenericTickListeners() As EnumerableCollection
Set GenericTickListeners = mGenericTickListeners
End Property

Public Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Get IsMarketDataRequested() As Boolean
IsMarketDataRequested = mIsMarketDataRequested
End Property

Public Property Get IsMarketDepthRequested() As Boolean
IsMarketDepthRequested = mIsMarketDepthRequested
End Property

Public Property Get IsTickReplay() As Boolean
IsTickReplay = False
End Property

Public Property Get Key() As String
Key = mKey
End Property

Friend Property Get MarketDepthListeners() As EnumerableCollection
checkNotFinished
Set MarketDepthListeners = mMarketDepthListeners
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
PriceChange = getDataStream.GetPriceChange()

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get PriceChangeListeners() As EnumerableCollection
Set PriceChangeListeners = mPriceChangeListeners
End Property

Friend Property Get QuoteListeners() As EnumerableCollection
Set QuoteListeners = mQuoteListeners
End Property

Friend Property Get RawMarketDepthListeners() As EnumerableCollection
Set RawMarketDepthListeners = mRawMarketDepthListeners
End Property

Public Property Get Session() As Session
Set Session = mSession
End Property

Public Property Get State() As MarketDataSourceStates
State = mState
End Property

Friend Property Get StreamId() As Long
Const ProcName As String = "StreamId"
On Error GoTo Err

StreamId = getDataStream.StreamId

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
Assert mState = MarketDataSourceStateRunning, "Not in Running state"
Timestamp = getDataStream.Timestamp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddErrorListener(ByVal pListener As ErrorListener)
Const ProcName As String = "AddErrorListener"
On Error GoTo Err

checkNotFinished
mErrorListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddGenericTickListener"
On Error GoTo Err

checkNotFinished
mGenericTickListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddMarketDepthListener(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "AddMarketDepthListener"
On Error GoTo Err

checkNotFinished
mMarketDepthListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddPriceChangeListener(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "AddPriceChangeListener"
On Error GoTo Err

checkNotFinished
mPriceChangeListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "AddQuoteListener"
On Error GoTo Err

checkNotFinished
mQuoteListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "AddRawMarketDepthListener"
On Error GoTo Err

checkNotFinished
mRawMarketDepthListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

checkNotFinished
mStateChangeListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ConvertTickToQuote(ByRef pTick As GenericTick) As Quote
Const ProcName As String = "ConvertTickToQuote"
On Error GoTo Err

checkNotFinished
ConvertTickToQuote = getDataStream.ConvertTickToQuote(pTick)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

checkNotFinished
getDataStream.RemoveMarketDataSource Me
setState MarketDataSourceStateFinished

Set mGenericTickListeners = Nothing
Set mMarketDepthListeners = Nothing
Set mPriceChangeListeners = Nothing
Set mQuoteListeners = Nothing
Set mRawMarketDepthListeners = Nothing
Set mStateChangeListeners = Nothing
Set mErrorListeners = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pKey As String, _
                ByVal pDataStream As MarketDataStream, _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "Initialise"
On Error GoTo Err

mKey = pKey
Set mDataStreamRef = CreateWeakReference(pDataStream)

setState MarketDataSourceStateCreated

Contract = pDataStream.Contract
Set mContractSpecifier = pDataStream.ContractSpecifier

Set mConfig = pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentDOM( _
                ByVal pListener As IMarketDepthListener)
Const ProcName As String = "NotifyCurrentDOM"
On Error GoTo Err

checkNotFinished
If mState = MarketDataSourceStateRunning Then getDataStream.DOMManager.NotifyCurrentDOM pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener)
Const ProcName As String = "NotifyCurrentMarketDataQuoteValues"
On Error GoTo Err

checkNotFinished
If mState = MarketDataSourceStateRunning Then getDataStream.NotifyCurrentMarketDataQuoteValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener)
Const ProcName As String = "NotifyCurrentMarketDataTickValues"
On Error GoTo Err

checkNotFinished
If mState = MarketDataSourceStateRunning Then getDataStream.NotifyCurrentMarketDataTickValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyCurrentMarketDepthValues( _
                ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "NotifyCurrentMarketDepthValues"
On Error GoTo Err

checkNotFinished
If mState = MarketDataSourceStateRunning Then getDataStream.DOMManager.NotifyCurrentMarketDepthValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Refresh()
Const ProcName As String = "Refresh"
On Error GoTo Err

checkNotFinished

If mContract Is Nothing Then Exit Sub

Dim lTickListener As IGenericTickListener
For Each lTickListener In mGenericTickListeners
    getDataStream.NotifyCurrentMarketDataTickValues lTickListener, Me
Next

Dim lQuoteListener As IQuoteListener
For Each lQuoteListener In mQuoteListeners
    getDataStream.NotifyCurrentMarketDataQuoteValues lQuoteListener, Me
Next

Dim lPricechangeListener As IPriceChangeListener
Dim ev As PriceChangeEventData
Set ev.Source = Me
ev.PriceChange = getDataStream.GetPriceChange
For Each lPricechangeListener In mPriceChangeListeners
    lPricechangeListener.Change ev
Next

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
ev.PriceChange = getDataStream.GetPriceChange
pListener.Change ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RefreshQuotes(ByVal pListener As IQuoteListener)
Const ProcName As String = "RefreshQuotes"
On Error GoTo Err

checkNotFinished
getDataStream.NotifyCurrentMarketDataQuoteValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RefreshTicks(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RefreshTicks"
On Error GoTo Err

checkNotFinished
getDataStream.NotifyCurrentMarketDataTickValues pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveErrorListener(ByVal pListener As ErrorListener)
Const ProcName As String = "RemoveErrorListener"
On Error GoTo Err

checkNotFinished
mErrorListeners.Remove pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveGenericTickListener"
On Error GoTo Err

checkNotFinished
mGenericTickListeners.Remove pListener

unsubscribeIfNoMarketDataListeners

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveMarketDepthListener(ByVal pListener As IMarketDepthListener)
Const ProcName As String = "RemoveMarketDepthListener"
On Error GoTo Err

checkNotFinished
mMarketDepthListeners.Remove pListener

unsubscribeIfNoMarketDepthListeners

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemovePriceChangeListener(ByVal pListener As IPriceChangeListener)
Const ProcName As String = "RemovePriceChangeListener"
On Error GoTo Err

checkNotFinished
mPriceChangeListeners.Remove pListener

unsubscribeIfNoMarketDataListeners

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveQuoteListener(ByVal pListener As IQuoteListener)
Const ProcName As String = "RemoveQuoteListener"
On Error GoTo Err

checkNotFinished
mQuoteListeners.Remove pListener

unsubscribeIfNoMarketDataListeners

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Const ProcName As String = "RemoveRawMarketDepthListener"
On Error GoTo Err

checkNotFinished
mRawMarketDepthListeners.Remove pListener

unsubscribeIfNoMarketDepthListeners

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

checkNotFinished
mStateChangeListeners.Remove pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMessage As String)
Const ProcName As String = "SetError"
On Error GoTo Err

Dim lListener As ErrorListener
Dim ev As ErrorEventData

Set ev.Source = Me
ev.ErrorCode = pErrorCode
ev.ErrorMessage = pErrorMessage

For Each lListener In mErrorListeners
    lListener.Notify ev
Next

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartMarketData()
Const ProcName As String = "StartMarketData"
On Error GoTo Err

checkNotFinished
Assert Not mIsMarketDataRequested, "Market data already started"

getDataStream.AddMarketDataSubscription Me

If Not mContract Is Nothing Then setState MarketDataSourceStateRunning
mIsMarketDataRequested = True

storeSettings

Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StartMarketDepth()
Const ProcName As String = "StartMarketDepth"
On Error GoTo Err

checkNotFinished
Assert Not mIsMarketDepthRequested, "Market depth already started"

getDataStream.AddMarketDepthSubscription Me

If Not mContract Is Nothing Then setState MarketDataSourceStateRunning
mIsMarketDepthRequested = True

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

getDataStream.RemoveMarketDataSubscription Me
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

getDataStream.RemoveMarketDepthSubscription Me
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
Assert mState <> MarketDataSourceStateFinished, "Object has been Finished"
End Sub

Private Sub fireStateChange(ByVal pState As MarketDataSourceStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim lListener As StateChangeListener
Dim ev As StateChangeEventData

Set ev.Source = Me
ev.State = pState

For Each lListener In mStateChangeListeners
    lListener.Change ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getDataStream() As MarketDataStream
Set getDataStream = mDataStreamRef.Target
End Function

Private Sub setState(ByVal pState As MarketDataSourceStates)
Const ProcName As String = "setState"
On Error GoTo Err

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

If Not mContract Is Nothing Then SaveContractToConfig mContract, mConfig.AddConfigurationSection(ConfigSectionContract)
mConfig.SetSetting ConfigSettingIsMarketDataRequested, mIsMarketDataRequested
mConfig.SetSetting ConfigSettingIsMarketDepthRequested, mIsMarketDepthRequested

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


