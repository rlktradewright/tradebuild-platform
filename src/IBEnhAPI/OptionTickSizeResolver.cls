VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionTickSizeResolver"
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

Implements ITask

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

Private Const ModuleName                            As String = "OptionTickSizeResolver"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mExamplePrice                               As Double
Private mOrderSubmitter                             As OrderSubmitter

Private mTickSize                                   As Double

Private mDummyOrder                                 As Order

Private mNewLimitPrice                              As Double

Private mFutureBuilder                              As New FutureBuilder

Private mTimerList                                  As TimerList
Private WithEvents mTLI                             As TimerListItem
Attribute mTLI.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mTickSize = 0.01
Set mTimerList = GetGlobalTimerList
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

mDummyOrder.DontCheckAttributeModificationPermitted = True
mDummyOrder.Action = OrderActionBuy
mDummyOrder.Quantity = 1
mDummyOrder.ErrorCode = 0
mDummyOrder.ErrorMessage = ""
mDummyOrder.DontCheckAttributeModificationPermitted = False
mOrderSubmitter.PlaceOrderEx mDummyOrder, mNewLimitPrice, False, 0, "", OrderModeEntry

Set mTLI = mTimerList.Add(Nothing, 500, ExpiryTimeUnitMilliseconds)

mTaskContext.Suspend -1


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
Const ProcName As String = "ITask_TaskName"
On Error GoTo Err

ITask_TaskName = mTaskContext.Name

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mTLI Event Handlers
'@================================================================================

Private Sub mTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mTLI_StateChange"
On Error GoTo Err

Dim lState As TimerListItemStates: lState = ev.State
If lState = TimerListItemStateExpired Then
    ' we haven't received an error during the timeout period, so the
    ' current tick size must be acceptable
    
    ' cancel the dummy order so that it is not transmitted when
    ' the 'real' order comes through
    mOrderSubmitter.CancelOrder mDummyOrder
    
    mTaskContext.Finish Nothing, False
    
    Dim lTickSize As Double
    If mExamplePrice < 3# Then
        lTickSize = mTickSize
    ElseIf mTickSize = 0.05 Then
        lTickSize = 0.01
    ElseIf mTickSize = 0.1 Then
        lTickSize = 0.05
    Else
        Assert False, "Unexpected tick size"
    End If
    mFutureBuilder.Value = CreateBoxedValue(lTickSize)
    mFutureBuilder.Complete
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get TickSizeFuture() As IFuture
Set TickSizeFuture = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pOrderSubmitter As OrderSubmitter, _
                ByVal pExamplePrice As Double, _
                ByVal pWantLowerPrice As Boolean, _
                ByVal pDummyOrder As Order)
mExamplePrice = deriveNonalignedPrice(pExamplePrice, pWantLowerPrice)
Set mOrderSubmitter = pOrderSubmitter
Set mDummyOrder = pDummyOrder

mNewLimitPrice = mExamplePrice
End Sub

Friend Function HandleError(ByVal pErrorNumber As Long) As Boolean
Const ProcName As String = "HandleError"
On Error GoTo Err

If pErrorNumber <> 110 Then Exit Function
mTLI.Cancel

' price does not conform to the minimum price variation. We need to
' adjust the limit price in the dummy order to use the next larger ticksize,
' and then resubmit the order.

Select Case mTickSize
Case 0.01
    mTickSize = 0.05
Case 0.05
    mTickSize = 0.1
Case 0.1
    mTickSize = 0.5
Case 0.5
    mTickSize = 1#
Case Else
    mTaskContext.Finish 0#, False
    HandleError = True
    Exit Function
End Select

mNewLimitPrice = adjustPrice(mExamplePrice, mTickSize)
mTaskContext.Continue

HandleError = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Public Function adjustPrice( _
                ByVal pPrice As Double, _
                ByVal pTicksize As Double) As Double
adjustPrice = Int((pPrice + (pTicksize / 2#) - 0.00001) / pTicksize) * pTicksize
End Function

Public Function deriveNonalignedPrice( _
                ByVal pPrice As Double, _
                ByVal pWantLowerPrice As Boolean) As Double
Dim lNewPrice As Double
Dim lIntegerPart As Double: lIntegerPart = Int(pPrice)
Dim lFractionalPart As Double: lFractionalPart = pPrice - lIntegerPart
If lFractionalPart < 0.5 Then
    lNewPrice = lIntegerPart + 0.37
    If pWantLowerPrice Then
        If lNewPrice > pPrice Then lNewPrice = lIntegerPart - 1 + 0.63
    Else
        If lNewPrice < pPrice Then lNewPrice = lIntegerPart + 0.63
    End If
Else
    lNewPrice = lIntegerPart + 0.63
    If pWantLowerPrice Then
        If lNewPrice > pPrice Then lNewPrice = lIntegerPart + 0.37
    Else
        If lNewPrice < pPrice Then lNewPrice = lIntegerPart + 1 + 0.37
    End If
End If
deriveNonalignedPrice = lNewPrice
End Function




