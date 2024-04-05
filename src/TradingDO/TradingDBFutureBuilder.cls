VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradingDBFutureBuilder"
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

Private Const ModuleName                            As String = "TradingDBFutureBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mFutureBuilder                              As New FutureBuilder
Private WithEvents mTradingDB                       As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mTradingDB Event Handlers
'@================================================================================

Private Sub mTradingDB_Connected()
Const ProcName As String = "mTradingDB_Connected"
On Error GoTo Err

mFutureBuilder.Value = mTradingDB
mFutureBuilder.Complete
Set mSelfRef = Nothing

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTradingDB_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
Const ProcName As String = "mTradingDB_ConnectFailed"
On Error GoTo Err

mFutureBuilder.Fail errorCode, errorDesc, ""
Set mSelfRef = Nothing

Exit Sub

Err:
GTDO.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Future() As IFuture
Set Future = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pConnectionParams As ConnectionParams, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me
Set mTradingDB = New TradingDB
mTradingDB.Initialise pConnectionParams
mFutureBuilder.cookie = pCookie

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




