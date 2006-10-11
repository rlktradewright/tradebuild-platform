VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tickers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================


Event error(ByRef ev As ErrorEvent)

Event stateChange(ByRef ev As StateChangeEvent)



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
'  Used to allocate handles for tickers. The handle is not used internally, but
'  can be used by the application to index quickly into tables or collections
'  of tickers. When a ticker is created, the first entry in this table that is
'  not set to true is allocated for the new ticker, and its index becomes the
'  ticker's handle.
'*/
Private mTickerHandleTable()        As Boolean

Private mTickers                    As Collection
Private mChangeListeners            As WeakCollection

Private mTradeBuildAPIRef           As WeakReference

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mTickerHandleTable(99) As Boolean
Set mTickers = New Collection
Set mChangeListeners = New WeakCollection

End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Count() As Long
Count = mTickers.Count
End Property

Public Property Get Item(ByVal index As Variant) As ticker
Attribute Item.VB_UserMemId = 0
Set Item = mTickers.Item(index)
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal ordersAreLive As Boolean, _
                Optional ByVal key As String) As ticker
Set add = New ticker
add.tradeBuildAPI = mTradeBuildAPIRef.target
add.handle = AllocateTickerHandle
add.Tickers = Me
add.simulatedOrders = Not ordersAreLive
If key = "" Then key = gGenerateGUIDString
add.key = key
mTickers.add add, key
fireChange CollItemAdded, add
End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Public Function createTickFileManager() As TickFileManager
Set createTickFileManager = New TickFileManager
createTickFileManager.Tickers = Me
createTickFileManager.tradeBuildAPI = mTradeBuildAPIRef.target
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mTickers.[_NewEnum]
End Function

Public Sub remove(ByVal pTicker As ticker)
If pTicker.State <> TickerStates.TickerStateClosing Then
    pTicker.StopTicker
    Exit Sub
End If
    
mTickers.remove pTicker.key
mTickerHandleTable(pTicker.handle) = False
fireChange CollItemRemoved, pTicker
End Sub

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.remove i
Next
End Sub

Friend Sub stateChange( _
                ByRef ev As StateChangeEvent)
RaiseEvent stateChange(ev)
End Sub

Friend Sub error(ByRef ev As ErrorEvent)
RaiseEvent error(ev)
End Sub

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = New WeakReference
mTradeBuildAPIRef.initialise value
End Property

'================================================================================
' Helper Functions
'================================================================================

Private Function AllocateTickerHandle() As Long
Dim i As Long

AllocateTickerHandle = -1

For i = 0 To UBound(mTickerHandleTable)
    If mTickerHandleTable(i) = False Then
        AllocateTickerHandle = i
        Exit For
    End If
Next

If AllocateTickerHandle = -1 Then
    AllocateTickerHandle = UBound(mTickerHandleTable) + 1
    ReDim Preserve mTickerHandleTable(UBound(mTickerHandleTable) + 100) As Boolean
End If

mTickerHandleTable(AllocateTickerHandle) = True

End Function

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As ticker)
Dim listener As CollectionChangeListener
Dim i As Long
Dim Change As CollectionChangeEvent
Set Change.source = Me
Change.changeType = changeType
Set Change.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change Change
Next
End Sub


