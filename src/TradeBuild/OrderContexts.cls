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

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEvent)

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

Private mOrderContexts As Collection
Private mChangeListeners As Collection

Private mEnabled As Boolean

Private mIsSimulated As Boolean

Private mTicker As Ticker
Private mPositionManager As PositionManager

Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter

Private mProviderReady As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mOrderContexts = New Collection
Set mChangeListeners = New Collection
mEnabled = True
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderContexts terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Count = mOrderContexts.Count
End Property

Friend Property Let enabled(ByVal value As Boolean)
Dim oc As OrderContext
mEnabled = value
For Each oc In mOrderContexts
    oc.enabled = mEnabled
Next
End Property

Public Property Get enabled() As Boolean
enabled = mEnabled
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
Dim oc As OrderContext
mIsSimulated = value
For Each oc In mOrderContexts
    oc.IsSimulated = mIsSimulated
Next
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Dim lOrderContext As OrderContext
Set mOrderSubmitter = value
For Each lOrderContext In mOrderContexts
    lOrderContext.orderSubmitter = value
Next
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

Friend Property Let Ticker(ByVal value As Ticker)
Dim lOrderContext As OrderContext
Set mTicker = value
For Each lOrderContext In mOrderContexts
    lOrderContext.Ticker = value
Next
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function add(ByVal name As String) As OrderContext
Set add = New OrderContext
add.Ticker = mTicker
add.PositionManager = mPositionManager
add.orderSubmitter = mOrderSubmitter
add.IsSimulated = mIsSimulated
add.name = name
mOrderContexts.add add, name
fireChange CollItemAdded, add
add.enabled = mEnabled
add.isReady = mProviderReady
End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Dim lOrderContext As OrderContext
Dim ev As CollectionChangeEvent

mChangeListeners.add value

' notify any existing items
Set ev.source = Me
ev.changeType = CollItemAdded
For Each lOrderContext In mOrderContexts
    Set ev.affectedItem = lOrderContext
    value.Change ev
Next
End Sub

Friend Sub finish()
Dim OrderContext As OrderContext

Set mTicker = Nothing
Set mPositionManager = Nothing
Set mOrderSubmitter = Nothing
Set mChangeListeners = Nothing

For Each OrderContext In mOrderContexts
    OrderContext.finish
Next
End Sub

Public Function item(index As Variant) As OrderContext
Attribute item.VB_UserMemId = 0
Set item = mOrderContexts(index)
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
    If mChangeListeners.item(i) Is value Then mChangeListeners.remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As OrderContext)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEvent
Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub


