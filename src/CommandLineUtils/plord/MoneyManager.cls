VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoneyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IMoneyManager

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

Private Const ModuleName                            As String = "MoneyManager"

'@================================================================================
' Member variables
'@================================================================================

Private mMaxOrderCost                               As Double
Private mLotSize                                     As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IMoneyManager Interface Members
'@================================================================================

Private Function IMoneyManager_AllowOrder( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pPositionManager As PositionManager, _
                ByVal pTimestamp As Date, _
                ByRef pNewOrderSize As Long, _
                ByRef pMessage As String, _
                ByRef pSessionMustBeEnded As Boolean) As Boolean
Const ProcName As String = "IMoneyManager_AllowOrder"
On Error GoTo Err

If mMaxOrderCost = 0 Then
    IMoneyManager_AllowOrder = True
    Exit Function
End If

Dim lContract As IContract: Set lContract = pPositionManager.ContractFuture.Value

Dim lPrice As Double: lPrice = IIf(pBracketOrder.LongPosition, _
                                pPositionManager.DataSource.CurrentTick(TickTypeAsk).Price, _
                                pPositionManager.DataSource.CurrentTick(TickTypeBid).Price)

pNewOrderSize = Int(mMaxOrderCost / (lContract.Specifier.Multiplier * lPrice))
If mLotSize <> 0 Then pNewOrderSize = mLotSize * Int(pNewOrderSize / mLotSize)

If pNewOrderSize > 0 Then
    pMessage = "Order quantity set to " & pNewOrderSize
    IMoneyManager_AllowOrder = True
Else
    pMessage = "Order quantity reduced to 0 due to lack of funds"
    IMoneyManager_AllowOrder = False
End If

If pMessage <> "" Then gWriteLineToConsole pBracketOrder.Key & ":" & pMessage, True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IMoneyManager_CloseAll( _
                ByVal pPositionManager As PositionManager, _
                ByVal pTimestamp As Date, _
                ByRef pMessage As String, _
                ByRef pSessionMustBeEnded As Boolean) As Boolean

End Function

Private Sub IMoneyManager_NotifyPositionFlat()

End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let MaxOrderCost(ByVal Value As Double)
mMaxOrderCost = Value
End Property

Public Property Let LotSize(ByVal Value As Long)
mLotSize = Value
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




