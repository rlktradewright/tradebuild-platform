VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimulatorOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "SimulatorOrder"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

''
' Indicates that the order has not been dealt with yet
'@/
Private mHandled                                As Boolean

''
' The underlying <code>order</code>
'@/
Private mOrder                                  As IOrder

''
' Indicates that the <code>order</code> is now active at the exchange
'@/
Private mAtExchange                             As Boolean

''
' The id of any parent <code>order</code>. If this is Nothing, then the
' <code>order</code> is not an attached order
'@/
Private mParentId                               As String

''
' If non-empty, indicates that the <code>order</code> belongs to an oca group
'@/
Private mOcaGroup                               As String

Private mTransmit                               As Boolean

''
' If true, indicates that the <code>order</code> is completed and can
' be deleted from the relevant collection
'@/
Private mIsFinished                             As Boolean

Private mListener                               As IOrderSubmissionListener

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let AtExchange(ByVal Value As Boolean)
mAtExchange = Value
End Property

Friend Property Get AtExchange() As Boolean
AtExchange = mAtExchange
End Property

Friend Property Let IsFinished(ByVal Value As Boolean)
mIsFinished = Value
End Property

Friend Property Get IsFinished() As Boolean
IsFinished = mIsFinished
End Property

Friend Property Let Handled(ByVal Value As Boolean)
mHandled = Value
End Property

Friend Property Get Handled() As Boolean
Handled = mHandled
End Property

Friend Function IsInSameOcaGroupAs(ByVal pOrder As SimulatorOrder) As Boolean
IsInSameOcaGroupAs = (mOrder.Id <> pOrder.Order.Id And _
                    mOcaGroup <> "" And _
                    mOcaGroup = pOrder.OcaGroup)
End Function

Friend Function IsParentOf(ByVal pOrder As SimulatorOrder) As Boolean
IsParentOf = (mOrder.Id = pOrder.ParentId)
End Function

Friend Function IsRelatedTo(ByVal pOrder As SimulatorOrder) As Boolean
If pOrder.Order.Id = mOrder.Id Then
    ' same order
    IsRelatedTo = False
Else
    IsRelatedTo = IsParentOf(pOrder) Or IsSiblingOf(pOrder) Or IsInSameOcaGroupAs(pOrder)
End If
End Function

Friend Function IsSiblingOf(ByVal pOrder As SimulatorOrder) As Boolean
IsSiblingOf = (mOrder.Id <> pOrder.Order.Id And _
                mParentId <> "" And _
                mParentId = pOrder.ParentId)
End Function

Friend Property Get Listener() As IOrderSubmissionListener
Set Listener = mListener
End Property

Friend Property Let OcaGroup(ByVal Value As String)
mOcaGroup = Value
End Property

Friend Property Get OcaGroup() As String
OcaGroup = mOcaGroup
End Property

Friend Property Let Order(ByVal Value As IOrder)
Set mOrder = Value
End Property

Friend Property Get Order() As IOrder
Set Order = mOrder
End Property

Friend Property Let ParentId(ByVal Value As String)
mParentId = Value
End Property

Friend Property Get ParentId() As String
ParentId = mParentId
End Property

Friend Property Let Transmit(ByVal Value As Boolean)
Const ProcName As String = "Transmit"
On Error GoTo Err

Assert Not (mTransmit And Not Value), "pTransmit cannot be set fron True to False"
mTransmit = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Transmit() As Boolean
Transmit = mTransmit
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pOrder As IOrder, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByVal pTransmit As Boolean, _
                ByVal pListener As IOrderSubmissionListener)
Set mOrder = pOrder
mParentId = pParentId
mOcaGroup = pOcaGroup
mTransmit = pTransmit
Set mListener = pListener
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

