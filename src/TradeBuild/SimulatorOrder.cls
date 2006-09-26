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

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

'/**
' The underlying <code>order</code>
'*/
Private mOrder As order

'/**
' Index into the simulator's MarketTable relevant to this <code>order</code>
'*/
Private mMarketIndex    As Long

'/**
' Indicates that the <code>order</code> is now active at the exchange
'*/
Private mAtExchange     As Boolean

'/**
' The id of any parent <code>order</code>. If this is Nothing, then the
' <code>order</code> is not an attached order
'*/
Private mParentId       As String

'/**
' If non-empty, indicates that the <code>order</code> belongs to an oca group
'*/
Private mOcaGroup       As String

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let atExchange(ByVal value As Boolean)
mAtExchange = value
End Property

Friend Property Get atExchange() As Boolean
atExchange = mAtExchange
End Property

Friend Property Let marketIndex(ByVal value As Long)
mMarketIndex = value
End Property

Friend Property Get marketIndex() As Long
marketIndex = mMarketIndex
End Property

Friend Property Let ocaGroup(ByVal value As String)
mOcaGroup = value
End Property

Friend Property Get ocaGroup() As String
ocaGroup = mOcaGroup
End Property

Friend Property Let order(ByVal value As order)
Set mOrder = value
End Property

Friend Property Get order() As order
Set order = mOrder
End Property

Friend Property Let parentId(ByVal value As String)
mParentId = value
End Property

Friend Property Get parentId() As String
parentId = mParentId
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

