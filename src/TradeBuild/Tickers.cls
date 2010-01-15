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

Event StateChange(ByRef ev As StateChangeEvent)

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

Private mTickers                    As Collection
Private mChangeListeners            As Collection

Private mWorkspaceRef               As WeakReference

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
Const ProcName As String = "Count"
Dim failpoint As Long
On Error GoTo Err

Count = mTickers.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Item(ByVal index As Variant) As Ticker
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As Long
On Error GoTo Err

Set Item = mTickers.Item(index)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let Workspace( _
                ByVal pWorkspace As Workspace)
Const ProcName As String = "Workspace"
Dim failpoint As Long
On Error GoTo Err

Set mWorkspaceRef = CreateWeakReference(pWorkspace)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Workspace() As Workspace
Const ProcName As String = "Workspace"
Dim failpoint As Long
On Error GoTo Err

Set Workspace = mWorkspaceRef.Target

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal pTickerOptions As TickerOptions, _
                Optional ByVal pKey As String) As Ticker
Const ProcName As String = "Add"
Dim failpoint As Long
On Error GoTo Err

If pKey = "" Then pKey = GenerateGUIDString
Set Add = New Ticker
Add.Initialise AllocateTickerHandle, _
                Me.Workspace, _
                pKey, _
                Me, _
                pTickerOptions
mTickers.Add Add, pKey
fireChange CollItemAdded, Add

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
Dim failpoint As Long
On Error GoTo Err

mChangeListeners.Add value, CStr(ObjPtr(value))

gNotifyExistingCollectionMembers mTickers, value, Me
Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function CreateTickFileManager( _
                ByVal pTickerOptions As Long) As TickFileManager
Const ProcName As String = "CreateTickFileManager"
Dim failpoint As Long
On Error GoTo Err

Set CreateTickFileManager = New TickFileManager
CreateTickFileManager.TickerOptions = pTickerOptions
CreateTickFileManager.Tickers = Me

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mTickers.[_NewEnum]
End Function

Public Function RefreshPriceChange( _
                ByVal listener As PriceChangeListener) As TaskController
Dim pcrt As New PriceChangeRefreshTask
Const ProcName As String = "RefreshPriceChange"
Dim failpoint As Long
On Error GoTo Err

If mTickers.Count = 0 Then Exit Function
pcrt.Initialise Me, listener
Set RefreshPriceChange = StartTask(pcrt, PriorityNormal)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function
                
Public Function RefreshQuotes( _
                ByVal listener As PriceChangeListener) As TaskController
Dim qrt As New QuotesRefreshTask
Const ProcName As String = "RefreshQuotes"
Dim failpoint As Long
On Error GoTo Err

If mTickers.Count = 0 Then Exit Function
qrt.Initialise Me, listener
Set RefreshQuotes = StartTask(qrt, PriorityNormal)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function
                
Public Sub Remove(ByVal pTicker As Ticker)
Dim i As Long

Const ProcName As String = "Remove"
Dim failpoint As Long
On Error GoTo Err

If pTicker.State <> TickerStates.TickerStateClosing Then
    pTicker.StopTicker
    Exit Sub
End If
    
For i = 1 To mTickers.Count
    If mTickers.Item(i) Is pTicker Then
        mTickers.Remove i
        Exit For
    End If
Next
mTickerHandleTable(pTicker.Handle) = False
fireChange CollItemRemoved, pTicker

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
Dim failpoint As Long
On Error GoTo Err

mChangeListeners.Remove CStr(ObjPtr(value))

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub StateChange( _
                ByRef ev As StateChangeEvent)
Const ProcName As String = "StateChange"
Dim failpoint As Long
On Error GoTo Err

RaiseEvent StateChange(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function StartTickers( _
                ByVal pTickerOptions As Long, _
                ByVal pContractSpec As ContractSpecifier, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String, _
                Optional ByVal interval As Long = 50) As TaskController
Dim lTickersStarter As New TickersStarter

Const ProcName As String = "StartTickers"
Dim failpoint As Long
On Error GoTo Err

lTickersStarter.Initialise Me, _
                            pTickerOptions, _
                            pDOMEventsRequired, _
                            poutputTickfileFormat, _
                            poutputTickfilePath, _
                            interval
Set StartTickers = StartTask(lTickersStarter, PriorityNormal, "StartTickers/" & pContractSpec.ToString, , pContractSpec)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function StartTickersFromContracts( _
                ByVal pTickerOptions As Long, _
                ByVal pContracts As Contracts, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String, _
                Optional ByVal interval As Long = 50) As TaskController
Dim lTickersStarter As New TickersStarter

Const ProcName As String = "StartTickersFromContracts"
Dim failpoint As Long
On Error GoTo Err

lTickersStarter.Initialise Me, _
                            pTickerOptions, _
                            pDOMEventsRequired, _
                            poutputTickfileFormat, _
                            poutputTickfilePath, _
                            interval
Set StartTickersFromContracts = StartTask(lTickersStarter, PriorityNormal, "StartTickersFromContracts/" & GenerateTextID, , pContracts)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AllocateTickerHandle() As Long
Dim i As Long

Const ProcName As String = "AllocateTickerHandle"
Dim failpoint As Long
On Error GoTo Err

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Ticker)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEvent
Const ProcName As String = "fireChange"
Dim failpoint As Long
On Error GoTo Err

Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub


