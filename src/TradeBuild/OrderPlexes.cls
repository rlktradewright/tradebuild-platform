VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPlexes"
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

Private mOrderPlexes As Collection
Private mChangeListeners As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mOrderPlexes = New Collection
Set mChangeListeners = New Collection
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
Count = mOrderPlexes.Count
End Property

'================================================================================
' Methods
'================================================================================

Friend Function add( _
                ByVal pOrderPlex As TradeBuild.orderPlex) As TradeBuild.orderPlex
Set add = pOrderPlex
mOrderPlexes.add pOrderPlex, pOrderPlex.key
fireChange CollItemAdded, pOrderPlex
End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Public Function Item( _
                index As Variant) As TradeBuild.orderPlex
Set Item = mOrderPlexes(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mOrderPlexes.[_NewEnum]
End Function

Public Sub remove( _
                ByVal pOrderPlex As orderPlex)
mOrderPlexes.remove pOrderPlex.key
fireChange CollItemRemoved, pOrderPlex
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
                ByVal affectedObject As orderPlex)
Dim listener As CollectionChangeListener
Dim Change As CollectionChangeEvent
Set Change.source = Me
Change.changeType = changeType
Set Change.affectedObject = affectedObject
For Each listener In mChangeListeners
    listener.Change Change
Next
End Sub





