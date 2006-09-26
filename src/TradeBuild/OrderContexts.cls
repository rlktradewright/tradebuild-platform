VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderContexts"
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
' Amendment history
'================================================================================
'
'
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

Private mOrderContexts As Collection
Private mChangeListeners As Collection

Private mEnabled As Boolean

Private mDummy As Boolean
Private mSimulated As Boolean

Private mTicker As Ticker
Private mPositionManager As PositionManager

Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter

Private mProviderReady As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mOrderContexts = New Collection
Set mChangeListeners = New Collection
mEnabled = True
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Count() As Long
Count = mOrderContexts.Count
End Property

Friend Property Let dummy(ByVal value As Boolean)
mDummy = value
End Property

Friend Property Let enabled(ByVal value As Boolean)
mEnabled = value
End Property

Public Property Get enabled() As Boolean
enabled = mEnabled
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Set mOrderSubmitter = value
End Property

Friend Property Let ProviderReady(ByVal value As Boolean)
Dim lOrderContext As OrderContext
mProviderReady = value
For Each lOrderContext In mOrderContexts
    lOrderContext.isReady = value
Next
End Property

Friend Property Let PositionManager(ByVal value As PositionManager)
Set mPositionManager = value
End Property

Friend Property Let simulated(ByVal value As Boolean)
mSimulated = value
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Set mTicker = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function add(ByVal name As String) As OrderContext
If Not mEnabled Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "TradeBuild.OrderContexts::add", _
                "Order context addition is disabled"
End If

Set add = New OrderContext
add.Ticker = mTicker
add.PositionManager = mPositionManager
add.orderSubmitter = mOrderSubmitter
add.dummy = mDummy
add.simulated = mSimulated
add.name = name
mOrderContexts.add add, name
fireChange CollItemAdded, add
add.isReady = mProviderReady
End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Friend Sub finish()
Dim OrderContext As OrderContext

Set mTicker = Nothing
Set mPositionManager = Nothing

For Each OrderContext In mOrderContexts
    OrderContext.finish
Next
End Sub

Public Function Item(index As Variant) As OrderContext
Attribute Item.VB_UserMemId = 0
Set Item = mOrderContexts(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mOrderContexts.[_NewEnum]
End Function

Public Sub remove( _
                ByVal pOrderContext As OrderContext)
mOrderContexts.remove pOrderContext.name
fireChange CollItemRemoved, pOrderContext
End Sub

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.remove i
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As OrderContext)
Dim listener As CollectionChangeListener
Dim Change As CollectionChangeEvent
Set Change.source = Me
Change.changeType = changeType
Set Change.affectedObject = affectedObject
For Each listener In mChangeListeners
    listener.Change Change
Next
End Sub


