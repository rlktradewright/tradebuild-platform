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

Private mEnabled As Boolean

Private mDummy As Boolean

Private mTicker As ticker
Private mPositionManager As positionManager

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mOrderContexts = New Collection
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

Friend Property Let positionManager(ByVal value As positionManager)
Set mPositionManager = value
End Property

Friend Property Let ticker(ByVal value As ticker)
Set mTicker = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function add(ByVal name As String) As OrderContext
If Not mEnabled Then
    err.Raise ErrorCodes.IllegalStateException, _
                "TradeBuild.OrderContexts::add", _
                "Order context addition is disabled"
End If

Set add = New OrderContext
add.ticker = mTicker
add.positionManager = mPositionManager
add.dummy = mDummy
add.name = name
mOrderContexts.add add, name
End Function

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

'================================================================================
' Helper Functions
'================================================================================


