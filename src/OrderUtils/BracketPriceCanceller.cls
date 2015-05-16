VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketPriceCanceller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IGenericTickListener

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

Private Const ModuleName                            As String = "BracketPriceCanceller"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelPrice                                As Double
Private mBracketOrder                               As BracketOrder
Private WithEvents mEntryOrder                      As Order
Attribute mEntryOrder.VB_VarHelpID = -1
Private mDataSource                                 As IMarketDataSource

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Select Case ev.Tick.TickType
Case TickTypeTrade
    
    Dim cancelPriceHit As Boolean
    
    If mBracketOrder.LongPosition Then
        Select Case mBracketOrder.EntryOrder.OrderType
        Case OrderTypes.OrderTypeLimit, _
                OrderTypes.OrderTypeLimitIfTouched, _
                OrderTypes.OrderTypeMarketIfTouched
            If ev.Tick.Price >= mCancelPrice Then cancelPriceHit = True
        Case OrderTypes.OrderTypeStop, _
            OrderTypes.OrderTypeStopLimit
            If ev.Tick.Price <= mCancelPrice Then cancelPriceHit = True
        End Select
    Else
        Select Case mBracketOrder.EntryOrder.OrderType
        Case OrderTypes.OrderTypeLimit, _
                OrderTypes.OrderTypeLimitIfTouched, _
                OrderTypes.OrderTypeMarketIfTouched
            If ev.Tick.Price <= mCancelPrice Then cancelPriceHit = True
        Case OrderTypes.OrderTypeStop, _
            OrderTypes.OrderTypeStopLimit
            If ev.Tick.Price >= mCancelPrice Then cancelPriceHit = True
        End Select
    End If
    
    If cancelPriceHit Then
        mBracketOrder.SelfCancel "Cancel price reached"
        Finish
    End If

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mEntryOrder Event Handlers
'@================================================================================

Private Sub mEntryOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mEntryOrder_Fill"
On Error GoTo Err

Finish

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mEntryOrder_StatusChanged()
Const ProcName As String = "mEntryOrder_StatusChanged"
On Error GoTo Err

If mEntryOrder.Status = OrderStatusCancelling Or _
    mEntryOrder.Status = OrderStatusCancelled Or _
    mEntryOrder.Status = OrderStatusRejected _
Then Finish

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pBracketOrder As BracketOrder, _
                ByVal pCancelPrice As Double, _
                ByVal pDataSource As IMarketDataSource)
Set mDataSource = pDataSource
If mDataSource Is Nothing Then Exit Sub
mDataSource.AddGenericTickListener Me

Set mBracketOrder = pBracketOrder
Set mEntryOrder = mBracketOrder.EntryOrder

mCancelPrice = pCancelPrice

Set mSelfRef = Me
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Finish()
mDataSource.RemoveGenericTickListener Me
Set mSelfRef = Nothing
End Sub


