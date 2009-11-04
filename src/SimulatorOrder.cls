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
Private mOrder As Order

''
' Index into the simulator's MarketTable relevant to this <code>order</code>
'@/
Private mMarketIndex    As Long

''
' Indicates that the <code>order</code> is now active at the exchange
'@/
Private mAtExchange     As Boolean

''
' The id of any parent <code>order</code>. If this is Nothing, then the
' <code>order</code> is not an attached order
'@/
Private mParentId       As String

''
' If non-empty, indicates that the <code>order</code> belongs to an oca group
'@/
Private mOcaGroup       As String

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

Friend Property Let MarketIndex(ByVal value As Long)
mMarketIndex = value
End Property

Friend Property Get MarketIndex() As Long
MarketIndex = mMarketIndex
End Property

Friend Property Let OcaGroup(ByVal value As String)
mOcaGroup = value
End Property

Friend Property Get OcaGroup() As String
OcaGroup = mOcaGroup
End Property

Friend Property Let Order(ByVal value As Order)
Set mOrder = value
End Property

Friend Property Get Order() As Order
Set Order = mOrder
End Property

Friend Property Let ParentId(ByVal value As String)
mParentId = value
End Property

Friend Property Get ParentId() As String
ParentId = mParentId
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

