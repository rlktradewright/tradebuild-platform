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
Implements TaskCompletionListener

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

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataRequester                        As MarketDataRequester
Private mTickerId                                   As Long

Private mContract                                   As Contract

Private mMarketDataListeners                        As EnumerableCollection
Private mMarketDepthListeners                       As EnumerableCollection

Private mErrorListeners                             As Collection

Private mContractRequester                          As ContractDetailsRequester

Private mTag                                        As Variant

Private mState                                      As MarketDataSourceStates

Private mHandle                                     As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

mHandle = gAllocateHandle
Set mErrorListeners = New Collection
Set mMarketDataListeners = New EnumerableCollection
Set mMarketDepthListeners = New EnumerableCollection
setState MarketDataSourceStateCreated

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub Class_Terminate()
gReleaseHandle mHandle
End Sub

'@================================================================================
' IMarketDataSource Interface Members
'@================================================================================

Private Sub IMarketDataSource_AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_AddGenericTickListener"
On Error GoTo Err

AddMarketDataListener pListener

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get IMarketDataSource_CurrentAskPrice() As Double
Const ProcName As String = "IMarketDataSource_CurrentAskPrice"
On Error GoTo Err

IMarketDataSource_CurrentAskPrice = CurrentAskPrice

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_CurrentAskSize() As Long
Const ProcName As String = "IMarketDataSource_CurrentAskSize"
On Error GoTo Err

IMarketDataSource_CurrentAskSize = CurrentAskSize

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_CurrentBidPrice() As Double
Const ProcName As String = "IMarketDataSource_CurrentBidPrice"
On Error GoTo Err

IMarketDataSource_CurrentBidPrice = CurrentBidPrice

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_CurrentBidSize() As Long
Const ProcName As String = "IMarketDataSource_CurrentBidSize"
On Error GoTo Err

IMarketDataSource_CurrentBidSize = CurrentBidSize

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_CurrentOpenInterest() As Long
Const ProcName As String = "IMarketDataSource_CurrentOpenInterest"
On Error GoTo Err

IMarketDataSource_CurrentOpenInterest = CurrentOpenInterest

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_CurrentTradePrice() As Double
Const ProcName As String = "IMarketDataSource_CurrentTradePrice"
On Error GoTo Err

IMarketDataSource_CurrentTradePrice = CurrentTradePrice

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_CurrentTradeSize() As Long
Const ProcName As String = "IMarketDataSource_CurrentTradeSize"
On Error GoTo Err

IMarketDataSource_CurrentTradeSize = CurrentTradeSize

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_CurrentVolume() As Long
Const ProcName As String = "IMarketDataSource_CurrentVolume"
On Error GoTo Err

IMarketDataSource_CurrentVolume = CurrentVolume

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_Handle() As Long
Const ProcName As String = "IMarketDataSource_Handle"
On Error GoTo Err

IMarketDataSource_Handle = Handle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_PreviousSessionClosePrice() As Double
Const ProcName As String = "IMarketDataSource_PreviousSessionClosePrice"
On Error GoTo Err

IMarketDataSource_PreviousSessionClosePrice = PreviousSessionClosePrice

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub IMarketDataSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_RemoveGenericTickListener"
On Error GoTo Err

RemoveMarketDataListener pListener

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get IMarketDataSource_SessionHighPrice() As Double
Const ProcName As String = "IMarketDataSource_SessionHighPrice"
On Error GoTo Err

IMarketDataSource_SessionHighPrice = SessionHighPrice

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_SessionLowPrice() As Double
Const ProcName As String = "IMarketDataSource_SessionLowPrice"
On Error GoTo Err

IMarketDataSource_SessionLowPrice = SessionLowPrice

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_State() As MarketDataSourceStates
Const ProcName As String = "IMarketDataSource_State"
On Error GoTo Err

IMarketDataSource_State = State

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IMarketDataSource_Timestamp() As Date
Const ProcName As String = "IMarketDataSource_Timestamp"
On Error GoTo Err

IMarketDataSource_Timestamp = Timestamp

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

Dim lContractDetails As TwsContractDetails

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
Else
    Set lContractDetails = ev.result
    setState MarketDataSourceStateRunning
    mTickerId = mMarketDataRequester.RequestMarketData(lContractDetails, Me, mMarketDataListeners)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Public Property Get CurrentAskPrice() As Double
Const ProcName As String = "CurrentAskPrice"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
CurrentAskPrice = mMarketDataRequester.CurrentAskPrice(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentAskSize() As Long
Const ProcName As String = "CurrentAskSize"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
CurrentAskSize = mMarketDataRequester.CurrentAskSize(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentBidPrice() As Double
Const ProcName As String = "CurrentBidPrice"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
CurrentBidPrice = mMarketDataRequester.CurrentBidPrice(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentBidSize() As Long
Const ProcName As String = "CurrentBidSize"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
CurrentBidSize = mMarketDataRequester.CurrentBidSize(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentOpenInterest() As Long
Const ProcName As String = "CurrentOpenInterest"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
CurrentOpenInterest = mMarketDataRequester.CurrentOpenInterest(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentTradePrice() As Double
Const ProcName As String = "CurrentTradePrice"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
CurrentTradePrice = mMarketDataRequester.CurrentTradePrice(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentTradeSize() As Long
Const ProcName As String = "CurrentTradeSize"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
CurrentTradeSize = mMarketDataRequester.CurrentTradeSize(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CurrentVolume() As Long
Const ProcName As String = "CurrentVolume"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
CurrentVolume = mMarketDataRequester.CurrentVolume(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Get PreviousSessionClosePrice() As Double
Const ProcName As String = "PreviousSessionClosePrice"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
PreviousSessionClosePrice = mMarketDataRequester.PreviousSessionClosePrice(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SessionHighPrice() As Double
Const ProcName As String = "SessionHighPrice"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
SessionHighPrice = mMarketDataRequester.SessionHighPrice(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SessionLowPrice() As Double
Const ProcName As String = "SessionLowPrice"
On Error GoTo Err

If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
SessionLowPrice = mMarketDataRequester.SessionLowPrice(mTickerId)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get State() As MarketDataSourceStates
State = mState
End Property

Public Property Let Tag(ByVal Value As Variant)
gSetVariant mTag, Value
End Property

Public Property Get Tag() As Variant
gSetVariant Tag, mTag
End Property

Public Property Get Timestamp() As Date
If mState <> MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not in Running state"
Timestamp = mMarketDataRequester.Timestamp(mTickerId)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddErrorListener(ByVal pListener As ErrorListener)
Const ProcName As String = "AddErrorListener"
On Error GoTo Err

mErrorListeners.Add pListener, gGetObjectKey(pListener)

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub AddMarketDataListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddMarketDataListener"
On Error GoTo Err

mMarketDataListeners.Add pListener, gGetObjectKey(pListener)
If mState = MarketDataSourceStateRunning Then mMarketDataRequester.AddMarketDataListener mTickerId, Me, pListener

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pMarketDataRequester As MarketDataRequester, _
                ByVal pContract As Contract, _
                ByVal pContractRequester As ContractDetailsRequester)
Set mMarketDataRequester = pMarketDataRequester
Set mContract = pContract
Set mContractRequester = pContractRequester
End Sub

Friend Sub NotifyMarketDataError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDataError"
On Error GoTo Err

fireError pErrorCode, pErrorMsg
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyMarketDepthError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

fireError pErrorCode, pErrorMsg
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveErrorListener(ByVal pListener As ErrorListener)
Const ProcName As String = "RemoveErrorListener"
On Error GoTo Err

mErrorListeners.Remove gGetObjectKey(pListener)

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveMarketDataListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveMarketDataListener"
On Error GoTo Err

If mState = MarketDataSourceStateRunning Then mMarketDataRequester.RemoveMarketDataListener mTickerId, Me, pListener
mMarketDataListeners.Remove gGetObjectKey(pListener)

If mMarketDataListeners.Count = 0 Then setState MarketDataSourceStateStopped
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub StartMarketData()
Const ProcName As String = "StartMarketData"
On Error GoTo Err

Dim lTask As ComboContractConversionTask

If State = MarketDataSourceStateStarting Or State = MarketDataSourceStateRunning Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Market data already started"
setState MarketDataSourceStateStarting

If mContract.Specifier.SecType = SecTypeCombo Then
    Set lTask = New ComboContractConversionTask
    lTask.Initialise mContractRequester, mContract.Specifier
    StartTask(lTask, PriorityNormal).AddTaskCompletionListener Me
Else
    setState MarketDataSourceStateRunning
    mTickerId = mMarketDataRequester.RequestMarketData(gContractToTwsContractDetails(mContract), Me, mMarketDataListeners)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub StartMarketDepth(Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "StartMarketDepth"
On Error GoTo Err

If mContract.Specifier.SecType = SecTypeCombo Then Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Market depth not supported for Combo contracts"

mTickerId = mMarketDataRequester.RequestMarketDepth(gContractToTwsContractDetails(mContract), Me, mMarketDepthListeners, pNumberOfRows)
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'Public Sub StopMarketData()
'Const ProcName As String = "StopMarketData"
'On Error GoTo Err
'
'If State = MarketDataSourceStateCreated Or State = MarketDataSourceStateStopped Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Market data not started"
'setState MarketDataSourceStateStopped
'
'mMarketDataRequester.RemoveMarketDataListener mTickerId, Me, mGenericTickListener
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Sub
'
'Public Sub StopMarketDepth()
'Const ProcName As String = "StopMarketDepth"
'On Error GoTo Err
'
'mMarketDataRequester.CancelMarketDepth mTickerId, Me, mGenericTickListener
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Sub
'
'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireError(ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "fireError"
On Error GoTo Err

Dim lListener As ErrorListener
Dim ev As ErrorEventData

Set ev.Source = Me
ev.ErrorCode = pErrorCode
ev.ErrorMessage = pErrorMessage

For Each lListener In mErrorListeners
    lListener.Notify ev
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setState(ByVal pState As MarketDataSourceStates)
mState = pState
End Sub



