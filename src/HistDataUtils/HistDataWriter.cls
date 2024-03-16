VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataWriter"
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

Private Const ModuleName                            As String = "HistDataWriter"

'@================================================================================
' Member variables
'@================================================================================

Private mTickStream                                 As TickStream
Private mContractFuture                             As IFuture

Private mOptions                                    As HistDataWriteOptions

Private WithEvents mTimeframeSetupTC                As TaskController
Attribute mTimeframeSetupTC.VB_VarHelpID = -1

Private mTimeframeHandlers                          As EnumerableCollection

Private mBarWriter                                  As IBarWriter

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mTimeframeSetupTC Event Handlers
'@================================================================================

Private Sub mTimeframeSetupTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTimeframeSetupTC_Completed"
On Error GoTo Err

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
Else
    Set mTimeframeHandlers = ev.Result
End If

Exit Sub

Err:
GHistData.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If mTimeframeHandlers Is Nothing Then
    If Not mTimeframeSetupTC Is Nothing Then mTimeframeSetupTC.CancelTask
    Exit Sub
End If
    
Dim s As String
If mContractFuture.IsAvailable Then
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    s = " (" & lContract.Specifier.LocalSymbol & ")"
End If

Dim lTimeframeHandler As TimeframeHandler
For Each lTimeframeHandler In mTimeframeHandlers
    lTimeframeHandler.Finish
    s = s & vbCrLf & "    " & lTimeframeHandler.TimePeriod.ToString & ":"
    If mOptions And HistDataWriteOptionWriteTradeBars Then
        s = s & " " & lTimeframeHandler.TradeBarsWritten & " trade bars written in " & lTimeframeHandler.TradeDatabaseWrites & " writes;"
    End If
    If mOptions And HistDataWriteOptionWriteBidAndAskBars Then
        s = s & " " & lTimeframeHandler.AskBarsWritten & " ask bars written in " & lTimeframeHandler.AskDatabaseWrites & " writes;"
        s = s & " " & lTimeframeHandler.BidBarsWritten & " bid bars written in " & lTimeframeHandler.BidDatabaseWrites & " writes;"
    End If
Next

Set mTickStream = Nothing

GHistData.Logger.Log "HistDataWriter finished" & s, ProcName, ModuleName

mTimeframeHandlers.Clear

If Not mBarWriter Is Nothing Then mBarWriter.Finish

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pOptions As HistDataWriteOptions, _
                ByVal pSaveIntervalSeconds As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pOptions And HistDataWriteOptionWriteTradeBars Then AssertArgument pHistDataStore.Supports(HistDataStoreCapabilityWriteTradeBars), "HistDataStore does not support writing trade bars"
If pOptions And HistDataWriteOptionWriteBidAndAskBars Then AssertArgument pHistDataStore.Supports(HistDataStoreCapabilityWriteBidAndAskBars), "HistDataStore does not support writing bid and ask bars"

GHistData.Logger.Log "Initialising HistDataWriter", ProcName, ModuleName, LogLevelHighDetail

Set mContractFuture = pContractFuture
mOptions = pOptions

Set mBarWriter = GHistDataUtils.CreateBufferedBarWriter(pHistDataStore, pOutputMonitor, mContractFuture)

Dim lTask As New TimeframeSetupTask
lTask.Initialise pHistDataStore, mBarWriter, pOutputMonitor, mContractFuture, pClockFuture, pStudyBase, pOptions, pSaveIntervalSeconds

Set mTimeframeSetupTC = StartTask(lTask, PriorityNormal)

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



