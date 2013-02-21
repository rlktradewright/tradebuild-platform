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

Implements Enumerable

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEventData)

Event StateChange(ByRef ev As StateChangeEventData)

'@================================================================================
' Constants
'@================================================================================


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
'  Used to allocate handles for Tickers. The Handle is not used internally, but
'  can be used by the application to index quickly into tables or collections
'  of Tickers. When a Ticker is created, the first entry in this table that is
'  not set to true is allocated for the new Ticker, and its index becomes the
'  Ticker's Handle. The Handle is only unique within workspace.
'@/
Private mTickerHandleTable()        As Boolean

Private WithEvents mTickers         As EnumerableCollection

Private mWorkspaceRef               As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTickerHandleTable(3) As Boolean
Set mTickers = New EnumerableCollection
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Const ProcName As String = "Enumerable_Enumerator"
On Error GoTo Err

Set Enumerable_Enumerator = GetEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function Enumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean

End Function

Private Sub Enumerable_Remove(pCookie As Variant)

End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean

End Function

Private Property Get Enumerable_UpdateSequence() As Long

End Property

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

Set ev.source = Me
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Const ProcName As String = "Count"

On Error GoTo Err

Count = mTickers.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let WorkSpace( _
                ByVal pWorkspace As WorkSpace)
Const ProcName As String = "Workspace"
On Error GoTo Err

Set mWorkspaceRef = CreateWeakReference(pWorkspace)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get WorkSpace() As WorkSpace
Const ProcName As String = "Workspace"
On Error GoTo Err

Set WorkSpace = mWorkspaceRef.Target

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal pTickerOptions As TickerOptions, _
                Optional ByVal pKey As String, _
                Optional ByVal pNumberOfMarketDepthRows As Long = 20) As Ticker
Const ProcName As String = "Add"
On Error GoTo Err

If pKey = "" Then pKey = GenerateGUIDString
Set Add = New Ticker
Add.Initialise AllocateTickerHandle, _
                Me.WorkSpace, _
                pKey, _
                Me, _
                pTickerOptions, _
                pNumberOfMarketDepthRows
mTickers.Add Add, pKey

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal pListener As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mTickers.AddCollectionChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function CreateTickFileManager( _
                ByVal pTickerOptions As Long) As TickFileManager
Const ProcName As String = "CreateTickFileManager"
On Error GoTo Err

Set CreateTickFileManager = New TickFileManager
CreateTickFileManager.TickerOptions = pTickerOptions
CreateTickFileManager.Tickers = Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item(ByVal pIndex As Variant) As Ticker
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mTickers.Item(pIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadTickersFromContracts( _
                ByVal pTickerOptions As Long, _
                ByVal pContracts As Contracts, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String, _
                Optional ByVal interval As Long = 50) As TaskController
Const ProcName As String = "LoadTickersFromContracts"
On Error GoTo Err

Dim lTickersStarter As New TickersStarter

lTickersStarter.Initialise Me, _
                            pTickerOptions, _
                            pDOMEventsRequired, _
                            poutputTickfileFormat, _
                            poutputTickfilePath, _
                            interval, _
                            True
Set LoadTickersFromContracts = StartTask(lTickersStarter, PriorityNormal, "LoadTickersFromContracts/" & GenerateTextID, , pContracts)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = GetEnumerator
End Function

Public Function RefreshPriceChange( _
                ByVal listener As PriceChangeListener) As TaskController
Const ProcName As String = "RefreshPriceChange"
On Error GoTo Err

Dim pcrt As New PriceChangeRefreshTask

If mTickers.Count = 0 Then Exit Function
pcrt.Initialise Me, listener
Set RefreshPriceChange = StartTask(pcrt, PriorityNormal)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                
Public Function RefreshQuotes( _
                ByVal listener As PriceChangeListener) As TaskController
Const ProcName As String = "RefreshQuotes"
On Error GoTo Err

Dim qrt As New QuotesRefreshTask

If mTickers.Count = 0 Then Exit Function
qrt.Initialise Me, listener
Set RefreshQuotes = StartTask(qrt, PriorityNormal)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                
Public Sub Remove(ByVal pTicker As Ticker)
Const ProcName As String = "Remove"
On Error GoTo Err

If pTicker.State <> TickerStates.TickerStateClosing Then
    pTicker.StopTicker
    Exit Sub
End If
    
mTickers.Remove pTicker.Key
mTickerHandleTable(pTicker.Handle) = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal pListener As CollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mTickers.RemoveCollectionChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StateChange( _
                ByRef ev As StateChangeEventData)
Const ProcName As String = "StateChange"
On Error GoTo Err

RaiseEvent StateChange(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartTickers( _
                ByVal pTickerOptions As Long, _
                ByVal pContractSpec As ContractSpecifier, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String, _
                Optional ByVal interval As Long = 50) As TaskController
Const ProcName As String = "StartTickers"
On Error GoTo Err

Dim lTickersStarter As New TickersStarter

lTickersStarter.Initialise Me, _
                            pTickerOptions, _
                            pDOMEventsRequired, _
                            poutputTickfileFormat, _
                            poutputTickfilePath, _
                            interval, _
                            False
Set StartTickers = StartTask(lTickersStarter, PriorityNormal, "StartTickers/" & pContractSpec.ToString, , pContractSpec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function StartTickersFromContracts( _
                ByVal pTickerOptions As Long, _
                ByVal pContracts As Contracts, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String, _
                Optional ByVal interval As Long = 50) As TaskController
Const ProcName As String = "StartTickersFromContracts"
On Error GoTo Err

Dim lTickersStarter As New TickersStarter

lTickersStarter.Initialise Me, _
                            pTickerOptions, _
                            pDOMEventsRequired, _
                            poutputTickfileFormat, _
                            poutputTickfilePath, _
                            interval, _
                            False
Set StartTickersFromContracts = StartTask(lTickersStarter, PriorityNormal, "StartTickersFromContracts/" & GenerateTextID, , pContracts)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AllocateTickerHandle() As Long
Const ProcName As String = "AllocateTickerHandle"
On Error GoTo Err

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

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function GetEnumerator() As Enumerator
Const ProcName As String = "GetEnumerator"
On Error GoTo Err

GetEnumerator = mTickers.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


