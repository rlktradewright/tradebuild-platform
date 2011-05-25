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

Private mGenericTickListeners                       As Collection
Private mNotificationListeners                      As Collection

Private mContractRequester                          As ContractDetailsRequester

Private mTag                                        As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mGenericTickListeners = New Collection
Set mNotificationListeners = New Collection
End Sub

'@================================================================================
' IMarketDataSource Interface Members
'@================================================================================

Private Sub IMarketDataSource_AddGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_AddGenericTickListener"
On Error GoTo Err

mGenericTickListeners.Add pListener, Hex$(ObjPtr(pListener))

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IMarketDataSource_AddMarketDepthListener(ByVal pListener As IMarketDepthListener)

End Sub

Private Sub IMarketDataSource_AddQuoteListener(ByVal pListener As IQuoteListener)

End Sub

Private Sub IMarketDataSource_AddRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)

End Sub

Private Sub IMarketDataSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
Const ProcName As String = "IMarketDataSource_RemoveGenericTickListener"
On Error GoTo Err

mGenericTickListeners.Remove Hex$(ObjPtr(pListener))

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IMarketDataSource_RemoveMarketDepthListener(ByVal pListener As IMarketDepthListener)

End Sub

Private Sub IMarketDataSource_RemoveQuoteListener(ByVal pListener As IQuoteListener)

End Sub

Private Sub IMarketDataSource_RemoveRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)

End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

Dim lContractDetails As TwsContract

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
Else
    Set lContractDetails = ev.result
    mTickerId = mMarketDataRequester.RequestMarketData(lContractDetails, Me)
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

Public Property Let Tag(ByVal Value As Variant)
gSetVariant mTag, Value
End Property

Public Property Get Tag() As Variant
gSetVariant Tag, mTag
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddNotificationListener(ByVal pListener As NotificationListener)
Const ProcName As String = "AddNotificationListener"
On Error GoTo Err

mNotificationListeners.Add pListener, Hex$(ObjPtr(pListener))

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
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

fireNotification pErrorCode, pErrorMsg
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyMarketDepthError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

fireNotification pErrorCode, pErrorMsg
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

fireGenericTick pTick

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveNotificationListener(ByVal pListener As NotificationListener)
Const ProcName As String = "IMarketDataSource_RemoveNotificationListener"
On Error GoTo Err

mNotificationListeners.Remove Hex$(ObjPtr(pListener))

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub StartMarketData()
Const ProcName As String = "StartMarketData"
On Error GoTo Err

Dim lTask As ComboContractConversionTask

If mContract.Specifier.Sectype = SecTypeCombo Then
    Set lTask = New ComboContractConversionTask
    lTask.Initialise mContractRequester, mContract.Specifier
    StartTask(lTask, PriorityNormal).AddTaskCompletionListener Me
Else
    mTickerId = mMarketDataRequester.RequestMarketData(gContractToTwsContractDetails(mContract), Me)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub StartMarketDepth(Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "StartMarketDepth"
On Error GoTo Err

If mContract.Specifier.Sectype = SecTypeCombo Then Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Market depth not supported for Combo contracts"

mMarketDataRequester.RequestMarketDepth gContractToTwsContractDetails(mContract), Me, pNumberOfRows
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub StopMarketData()
Const ProcName As String = "StopMarketData"
On Error GoTo Err

mMarketDataRequester.CancelMarketData mTickerId, Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub StopMarketDepth()
Const ProcName As String = "StopMarketDepth"
On Error GoTo Err

mMarketDataRequester.CancelMarketDepth mTickerId, Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireGenericTick(ByRef pTick As GenericTick)
Const ProcName As String = "fireGenericTick"
On Error GoTo Err

Dim lListener As IGenericTickListener
Dim ev As GenericTickEventData

Set ev.source = Me
ev.Tick = pTick

For Each lListener In mGenericTickListeners
    lListener.NotifyTick ev
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub fireNotification(ByVal pEventCode As Long, ByVal pEventMessage As String)
Const ProcName As String = "fireNotification"
On Error GoTo Err

Dim lListener As NotificationListener
Dim ev As NotificationEventData

Set ev.source = Me
ev.eventCode = pEventCode
ev.eventMessage = pEventMessage

For Each lListener In mNotificationListeners
    lListener.Notify ev
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub



