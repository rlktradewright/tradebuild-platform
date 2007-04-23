VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradingDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event ConnectSucceeded()
Event ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "TradingDB"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mConnection As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1

Private mExchangeFactory As exchangeFactory
Private mInstrumentFactory As instrumentFactory

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mInstrumentFactory = New instrumentFactory
mInstrumentFactory.tradingDB = Me

Set mExchangeFactory = New exchangeFactory
mExchangeFactory.tradingDB = Me
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@===============================================================================
' mConnection Event Handlers
'@===============================================================================

Private Sub mConnection_ConnectComplete( _
                            ByVal pError As ADODB.Error, _
                            adStatus As ADODB.EventStatusEnum, _
                            ByVal pConnection As ADODB.Connection)

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    RaiseEvent ConnectFailed(pError.Number, _
                            ModuleName & ": connection error" & vbCrLf & _
                            pError.Description & vbCrLf & _
                            gGenerateErrorMessage(pConnection))
Else
    RaiseEvent ConnectSucceeded
    mConnection.Close
End If


End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get instrumentFactory() As instrumentFactory
Set instrumentFactory = mInstrumentFactory
End Property

Public Property Get exchangeFactory() As exchangeFactory
Set exchangeFactory = mExchangeFactory
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal connectionString As String)
mInstrumentFactory.connectionString = connectionString
mExchangeFactory.connectionString = connectionString

Set mConnection = New ADODB.Connection
mConnection.CursorLocation = CursorLocationEnum.adUseClient
mConnection.Open connectionString, , , ConnectOptionEnum.adAsyncConnect
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


