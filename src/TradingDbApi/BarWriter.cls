VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements IBarWriter
Implements TaskCompletionListener

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

Private Type TimeframeEntry
    Timeframe               As TimePeriod
    PrevTimestamp           As Date
    BarsWritten             As Long
    Writes                  As Long
End Type

'@===============================================================================
' Constants
'@===============================================================================

Private Const ModuleName                    As String = "BarWriter"

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                          As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1

Private mParent                             As HistDataStore
Private mOutputMonitor                      As IBarOutputMonitor

Private mReady                              As Boolean

Private mContractFuture                     As IFuture
Private mContract                           As IContract

Private mPath                               As String

Private mInstrument                         As Instrument

Private mTimeframeTable()                   As TimeframeEntry

Private mUseSynchronousReads                As Boolean
Private mUseSynchronousWrites               As Boolean

Private WithEvents mInstrumentTC            As TaskController
Attribute mInstrumentTC.VB_VarHelpID = -1

Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@===============================================================================
' IBarWriter Interface Members
'@===============================================================================

Private Property Get IBarWriter_ContractFuture() As IFuture
Set IBarWriter_ContractFuture = mContractFuture
End Property

Private Sub IBarWriter_Finish()
Const ProcName As String = "IBarWriter_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IBarWriter_IsReady() As Boolean
IBarWriter_IsReady = mReady
End Property

Private Sub IBarWriter_WriteBar( _
                ByVal pBar As Bar, _
                ByVal pBarType As BarTypes, _
                ByVal pTimePeriod As TimePeriod)
Const ProcName As String = "IBarWriter_WriteBar"
On Error GoTo Err

With pBar
    If mUseSynchronousWrites Then
        mInstrument.WriteBar .TimeStamp, _
                            pBarType, _
                            pTimePeriod.Length, _
                            .OpenValue, _
                            .HighValue, _
                            .LowValue, _
                            .CloseValue, _
                            .Volume, _
                            .TickVolume, _
                            .OpenInterest
    Else
        mInstrument.WriteBarAsync(.TimeStamp, _
                                pBarType, _
                                pTimePeriod.Length, _
                                .OpenValue, _
                                .HighValue, _
                                .LowValue, _
                                .CloseValue, _
                                .Volume, _
                                .TickVolume, _
                                .OpenInterest).AddTaskCompletionListener Me
    End If
    
    Dim index As Long
    For index = 0 To UBound(mTimeframeTable)
        If mTimeframeTable(index).Timeframe Is pTimePeriod Then
            With mTimeframeTable(index)
                If .PrevTimestamp <> pBar.TimeStamp Then
                    .BarsWritten = .BarsWritten + 1
                    .PrevTimestamp = pBar.TimeStamp
                End If
                .Writes = .Writes + 1
            End With
            Exit For
        End If
    Next
    
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarWriter_WriteComment(ByVal pComment As String, ByVal pTimestamp As Date)

End Sub

'@===============================================================================
' TaskCompletionListener Interface Members
'@===============================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_taskCompleted"
On Error GoTo Err

Dim lEv As NotificationEventData
If ev.ErrorNumber <> 0 Then
    Dim lMessage As String
    lMessage = "Error (" & ev.ErrorNumber & "): " & ev.ErrorMessage
    logMessage lMessage, ProcName, , LogLevelHighDetail
    
    Set lEv.Source = Me
    lEv.EventCode = ev.ErrorNumber
    lEv.EventMessage = ev.ErrorMessage
    If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyEvent lEv
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

If gLogger.IsLoggable(LogLevelDetail) Then logMessage pMsg:="Retrieving instrument object for " & mContract.Specifier.ToString, pProcName:="initialise", pLogLevel:=LogLevelDetail
Set mInstrumentTC = mTradingDB.InstrumentFactory.LoadBySpecifierAsync(mContract.Specifier)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@===============================================================================
' mInstrumentTC Event Handlers
'@===============================================================================

Private Sub mInstrumentTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mInstrumentTC_Completed"
On Error GoTo Err

Dim lEv As NotificationEventData
Set lEv.Source = Me

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    lEv.EventCode = ev.ErrorNumber
    lEv.EventMessage = ev.ErrorMessage
    If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyEvent lEv
ElseIf IsEmpty(ev.Result) Then
    lEv.EventCode = HistDataEventCodes.HistDataEventNoContractDetails
    lEv.EventMessage = "Error: specified instrument does not exist: " & mContract.Specifier.ToString
    If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyEvent lEv
Else
    Set mInstrument = ev.Result
    
    setupTimeframeTable
    
    mReady = True
    If Not mOutputMonitor Is Nothing Then
        mOutputMonitor.NotifyReady Me
        mOutputMonitor.NotifyOutputFileCreated Me, mContract.Specifier.ToString & ": " & mTradingDB.ToString
    End If
End If
Set mInstrumentTC = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@===============================================================================
' Properties
'@===============================================================================

Friend Property Let Path(ByVal Value As String)
mPath = Value
End Property

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

Dim s As String
s = "Bar writer (" & mContract.Specifier.LocalSymbol & ")"

Dim i As Long
For i = 0 To UBound(mTimeframeTable)
    s = s & vbCrLf & "    Total of " & mTimeframeTable(i).BarsWritten & " bars (" & _
            mTimeframeTable(i).Timeframe.ToString & _
            ") written in " & _
            mTimeframeTable(i).Writes & " database writes"
Next
logMessage s, ProcName

mParent.BarWriterFinished Me
Set mParent = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTradingDBFuture As IFuture, _
                ByVal pParent As HistDataStore, _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                ByVal pUseSynchronousReads As Boolean, _
                ByVal pUseSynchronousWrites As Boolean)
Const ProcName As String = "initialise"
On Error GoTo Err

If pTradingDBFuture.IsAvailable Then
    Set mTradingDB = pTradingDBFuture.Value
Else
    mFutureWaiter.Add pTradingDBFuture
End If

Set mParent = pParent
Set mOutputMonitor = pOutputMonitor

mUseSynchronousReads = pUseSynchronousReads
mUseSynchronousWrites = pUseSynchronousWrites

Set mContractFuture = pContractFuture
mFutureWaiter.Add mContractFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLogger.Log pMsg, pProcName, ModuleName, pLogLevel, pMsgQualifier
End Sub

Private Sub setupTimeframeTable()
Const ProcName As String = "requestTimeframes"
On Error GoTo Err

Dim lTimeperiods() As TimePeriod
lTimeperiods = mParent.WriterTimeframes

ReDim mTimeframeTable(UBound(lTimeperiods)) As TimeframeEntry

Dim i As Long
For i = 0 To UBound(mTimeframeTable)
    Set mTimeframeTable(i).Timeframe = lTimeperiods(i)
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

