VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DBClient"

'@================================================================================
' Member variables
'@================================================================================

Private mTradingDBFuture                            As IFuture
Private mUseSynchronousReads                        As Boolean
Private mUseSynchronousWrites                       As Boolean

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

Public Property Get ContractStore() As IContractStore
Const ProcName As String = "ContractStore"
On Error GoTo Err

Static sContractStore As ContractStore
If sContractStore Is Nothing Then
    Set sContractStore = New ContractStore
    sContractStore.Initialise mTradingDBFuture
End If
Set ContractStore = sContractStore

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HistoricalDataStore() As IHistoricalDataStore
Const ProcName As String = "HistDataStore"
On Error GoTo Err

Static sHistDataStore As HistDataStore
If sHistDataStore Is Nothing Then
    Set sHistDataStore = New HistDataStore
    sHistDataStore.Initialise mTradingDBFuture, mUseSynchronousReads, mUseSynchronousWrites
End If
Set HistoricalDataStore = sHistDataStore

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickfileStore() As ITickfileStore
Const ProcName As String = "TickfileStore"
On Error GoTo Err

Static sTickfileStore As TickfileStore
If sTickfileStore Is Nothing Then
    Set sTickfileStore = New TickfileStore
    sTickfileStore.Initialise mTradingDBFuture, mUseSynchronousReads, mUseSynchronousWrites
End If
Set TickfileStore = sTickfileStore

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TradingDBFuture() As IFuture
Set TradingDBFuture = mTradingDBFuture
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTradingDBFuture As IFuture, _
                ByVal pUseSynchronousReads As Boolean, _
                ByVal pUseSynchronousWrites As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTradingDBFuture = pTradingDBFuture
mUseSynchronousReads = pUseSynchronousReads
mUseSynchronousWrites = pUseSynchronousWrites

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




