VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@===============================================================================
' Description
'@===============================================================================

'@===============================================================================
' Interfaces
'@===============================================================================

Implements ITickfileWriter

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================


Private Const ModuleName As String = "TickfileWriter"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                  As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1

Private mParent                     As TickfileStore
Private mOutputMonitor              As ITickfileOutputMonitor

Private mContractFuture             As IFuture
Private mContract                   As IContract
Private WithEvents mTickWriter      As TickWriter
Attribute mTickWriter.VB_VarHelpID = -1

Private mTickSize                   As Double

Private mUseSynchronousWrites       As Boolean

Private mInstrument                 As Instrument
Private WithEvents mInstrumentTC    As TaskController
Attribute mInstrumentTC.VB_VarHelpID = -1

Private WithEvents mFutureWaiter    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mReady                      As Boolean

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@===============================================================================
' ITickfileWriter Interface Members
'@===============================================================================

Private Property Get ITickfileWriter_ContractFuture() As IFuture
Set ITickfileWriter_ContractFuture = mContractFuture
End Property

Private Sub ITickfileWriter_Finish()
Const ProcName As String = "ITickfileWriter_Finish"
On Error GoTo Err

mParent.TickfileWriterFinished Me
Finish

Exit Sub
Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickfileWriter_IsReady() As Boolean
ITickfileWriter_IsReady = mReady
End Property

Private Sub ITickfileWriter_WriteComment(ByVal pComment As String, ByVal pTimestamp As Date)

End Sub

Private Sub ITickfileWriter_WriteTick(pTick As GenericTick)
Const ProcName As String = "ITickfileWriter_WriteTick"
On Error GoTo Err

mTickWriter.WriteTick pTick

Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' mFutureWaiter Event Handlers
'@===============================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is TradingDB Then
    Set mTradingDB = ev.Future.Value
Else
    Set mContract = ev.Future.Value
End If

If mTradingDB Is Nothing Or mContract Is Nothing Then Exit Sub

If GTradingDB.Logger.IsLoggable(LogLevelDetail) Then
    logMessage "Retrieving instrument object for " & mContract.Specifier.ToString, ProcName, , LogLevelDetail
End If
Set mInstrumentTC = mTradingDB.InstrumentFactory.LoadBySpecifierAsync(mContract.Specifier)

Exit Sub

Err:
GTradingDB.NotifyUnhandledError ProcName, ModuleName
End Sub

'@===============================================================================
' mInstrumentTC Event Handlers
'@===============================================================================

Private Sub mInstrumentTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mInstrumentTC_Completed"
On Error GoTo Err

Dim lEv As NotificationEventData
Set lEv.Source = mParent

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    lEv.EventCode = TickfileEventCodes.TickfileEventDataSourceNotAvailable
    lEv.EventMessage = ev.ErrorMessage
    If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyEvent lEv
ElseIf IsEmpty(ev.Result) Then
    lEv.EventCode = TickfileEventCodes.TickfileEventNoContractDetails
    lEv.EventMessage = "Specified instrument does not exist: " & mContract.Specifier.ToString
    If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyEvent lEv
Else
    Set mInstrument = ev.Result
    mTickSize = mInstrument.TickSize
    Set mTickWriter = mInstrument.CreateTickWriter(Not mUseSynchronousWrites)
    
    mReady = True
    If Not mOutputMonitor Is Nothing Then
        mOutputMonitor.NotifyOutputFileCreated Me, mContract.Specifier.ToString & ": " & mTradingDB.ToString
        mOutputMonitor.NotifyReady Me
    End If
End If

Set mInstrumentTC = Nothing

Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' mTickWriter Event Handlers
'@===============================================================================

Private Sub mTickWriter_ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
Const ProcName As String = "mTickWriter_ConnectFailed"
On Error GoTo Err

Dim lEv As NotificationEventData
Set lEv.Source = Me

lEv.EventCode = TickfileEventCodes.TickfileEventDataSourceNotAvailable
lEv.EventMessage = "Connect error " & errorCode & ": " & errorDesc
If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyEvent lEv

Exit Sub
Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mTickWriter_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
Const ProcName As String = "mTickWriter_Error"
On Error GoTo Err

Dim lEv As NotificationEventData
Set lEv.Source = Me

lEv.EventCode = TickfileEventCodes.TickfileEventDataSourceNotAvailable
lEv.EventMessage = "Write error " & errorCode & ": " & errorDesc
If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyEvent lEv

' recreate the tickwriter as it seems to get buggered at this point
logMessage "Re-creating tickfile writer", ProcName
Set mTickWriter = mInstrument.CreateTickWriter(Not mUseSynchronousWrites)

Exit Sub
Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mInstrumentTC Is Nothing Then
    mInstrumentTC.CancelTask
    Set mInstrumentTC = Nothing
End If

If Not mTickWriter Is Nothing Then
    mTickWriter.Finish
    
    logMessage "Tickfile writer (" & mContract.Specifier.LocalSymbol & ")" & vbCrLf & _
                        "    Total data written: " & mTickWriter.BytesWritten & " bytes", _
                ProcName
    Set mTickWriter = Nothing
End If

Set mParent = Nothing

Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTradingDBFuture As IFuture, _
                ByVal pTickfileOutputMonitor As ITickfileOutputMonitor, _
                ByVal pParent As TickfileStore, _
                ByVal pContractFuture As IFuture, _
                ByVal pFormatIdentifier As String, _
                ByVal pLocation As String, _
                ByVal pUseSynchronousWrites As Boolean)
Const ProcName As String = "initialise"
On Error GoTo Err

If pTradingDBFuture.IsAvailable Then
    Set mTradingDB = pTradingDBFuture.Value
Else
    mFutureWaiter.Add pTradingDBFuture
End If

Set mOutputMonitor = pTickfileOutputMonitor
Set mParent = pParent

mUseSynchronousWrites = pUseSynchronousWrites

Set mContractFuture = pContractFuture
mFutureWaiter.Add pContractFuture

Exit Sub

Err:
GTradingDB.HandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
GTradingDB.Logger.Log pMsg, pProcName, ModuleName, pLogLevel, pMsgQualifier
End Sub




