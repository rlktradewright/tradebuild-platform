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

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEvent)

Event Notification(ByRef ev As NotificationEvent)

Event stateChange(ByRef ev As StateChangeEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TradeBuild26"
Private Const ModuleName                    As String = "Tickers"

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
'  Used to allocate handles for tickers. The handle is not used internally, but
'  can be used by the application to index quickly into tables or collections
'  of tickers. When a ticker is created, the first entry in this table that is
'  not set to true is allocated for the new ticker, and its index becomes the
'  ticker's handle. The handle is only unique within workspace.
'@/
Private mTickerHandleTable()        As Boolean

Private mTickers                    As Collection
Private mChangeListeners            As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTickerHandleTable(3) As Boolean
Set mTickers = New Collection
Set mChangeListeners = New Collection

End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Count = mTickers.Count
End Property

Public Property Get item(ByVal index As Variant) As Ticker
Attribute item.VB_UserMemId = 0
Set item = mTickers.item(index)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function add( _
                ByVal pTickerOptions As TickerOptions, _
                Optional ByVal key As String) As Ticker
If key = "" Then key = GenerateGUIDString
Set add = New Ticker
add.key = key
add.initialise AllocateTickerHandle, _
                Me, _
                pTickerOptions
mTickers.add add, key
fireChange CollItemAdded, add
End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Public Function createTickFileManager( _
                ByVal pTickerOptions As Long) As TickFileManager
Set createTickFileManager = New TickFileManager
createTickFileManager.TickerOptions = pTickerOptions
createTickFileManager.Tickers = Me
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mTickers.[_NewEnum]
End Function

Public Sub remove(ByVal pTicker As Ticker)
Dim i As Long

If pTicker.State <> TickerStates.TickerStateClosing Then
    pTicker.stopTicker
    Exit Sub
End If
    
For i = 1 To mTickers.Count
    If mTickers.item(i) Is pTicker Then
        mTickers.remove i
        Exit For
    End If
Next
mTickerHandleTable(pTicker.handle) = False
fireChange CollItemRemoved, pTicker
End Sub

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.remove i
Next
End Sub

Friend Sub stateChange( _
                ByRef ev As StateChangeEvent)
RaiseEvent stateChange(ev)
End Sub

Public Function startTickers( _
                ByVal pTickerOptions As Long, _
                ByVal pContractSpec As contractSpecifier, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String, _
                Optional ByVal interval As Long = 50) As TaskController
Dim lTickersStarter As New TickersStarter

Dim failpoint As Long
On Error GoTo Err

lTickersStarter.initialise Me, _
                            pTickerOptions, _
                            pDOMEventsRequired, _
                            poutputTickfileFormat, _
                            poutputTickfilePath, _
                            interval
Set startTickers = StartTask(lTickersStarter, PriorityNormal, "StartTickers/" & pContractSpec.toString, , pContractSpec)

Exit Function

Err:
Err.Raise Err.number, _
        ProjectName & "." & ModuleName & ":" & "startTickers" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Function

Public Function startTickersFromContracts( _
                ByVal pTickerOptions As Long, _
                ByVal pContracts As Contracts, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String, _
                Optional ByVal interval As Long = 50) As TaskController
Dim lTickersStarter As New TickersStarter

Dim failpoint As Long
On Error GoTo Err

lTickersStarter.initialise Me, _
                            pTickerOptions, _
                            pDOMEventsRequired, _
                            poutputTickfileFormat, _
                            poutputTickfilePath, _
                            interval
Set startTickersFromContracts = StartTask(lTickersStarter, PriorityNormal, "StartTickersFromContracts/" & GenerateTextID, , pContracts)

Exit Function

Err:
Err.Raise Err.number, _
        ProjectName & "." & ModuleName & ":" & "startTickersFromContracts" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

End Function

Friend Sub Notify(ByRef ev As NotificationEvent)
RaiseEvent Notification(ev)
gTB.Notify ev
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

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
    ReDim Preserve mTickerHandleTable(2 * (UBound(mTickerHandleTable) + 1) - 1) As Boolean
End If

mTickerHandleTable(AllocateTickerHandle) = True

End Function

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Ticker)
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


