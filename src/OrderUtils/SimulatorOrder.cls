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
' The underlying <code>order</code>
'@/
Private mOrder                                  As IOrder

''
' Index into the simulator's MarketTable relevant to this <code>order</code>
'@/
Private mMarketId                               As Long

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

Private mClientId                               As Long

Private mSimulatorRef                           As WeakReference

Private mTransmit                               As Boolean

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

Friend Property Let AtExchange(ByVal value As Boolean)
mAtExchange = value
End Property

Friend Property Get AtExchange() As Boolean
AtExchange = mAtExchange
End Property

Friend Property Let ClientId(ByVal value As Long)
mClientId = value
End Property

Friend Property Get ClientId() As Long
ClientId = mClientId
End Property

Friend Property Get MarketId() As Long
MarketId = mMarketId
End Property

Friend Property Let OcaGroup(ByVal value As String)
mOcaGroup = value
End Property

Friend Property Get OcaGroup() As String
OcaGroup = mOcaGroup
End Property

Friend Property Let Order(ByVal value As IOrder)
Set mOrder = value
End Property

Friend Property Get Order() As IOrder
Set Order = mOrder
End Property

Friend Property Let ParentId(ByVal value As String)
mParentId = value
End Property

Friend Property Get ParentId() As String
ParentId = mParentId
End Property

Friend Property Let Transmit(ByVal value As Boolean)
Const ProcName As String = "Transmit"
On Error GoTo Err

Assert Not (mTransmit And Not value), "pTransmit cannot be set fron True to False"
mTransmit = value

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
                ByVal pSimulator As InProcessExchangeSimulator, _
                ByVal pOrder As IOrder, _
                ByVal pClientId As Long, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByVal pTransmit As Boolean, _
                ByVal pMarketId As Long)
Set mSimulatorRef = CreateWeakReference(pSimulator)
Set mOrder = pOrder
mClientId = pClientId
mParentId = pParentId
mOcaGroup = pOcaGroup
mTransmit = pTransmit
mMarketId = pMarketId
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getSimulator() As InProcessExchangeSimulator
Set getSimulator = mSimulatorRef.Target
End Function
