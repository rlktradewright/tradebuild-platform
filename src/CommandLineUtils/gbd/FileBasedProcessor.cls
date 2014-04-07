VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileBasedProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Private Const ModuleName                    As String = "FileBasedProcessor"

'@================================================================================
' Member variables
'@================================================================================

Private mSelfRef As Object

Private WithEvents mReplayController As ReplayController
Attribute mReplayController.VB_VarHelpID = -1

Private WithEvents mReplayTaskController As TaskController
Attribute mReplayTaskController.VB_VarHelpID = -1

Private mNumber As Long
Private mBarLength As Long
Private mBarUnits As TimePeriodUnits
Private mSessionOnly As Boolean

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mTB As TradeBuildAPI

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' mReplayTaskController Event Handlers
'@================================================================================

Private Sub mReplayTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mReplayTaskController_Completed"
On Error GoTo Err

gProcessor.Finish
Set gProcessor = Nothing
gCon.WriteLineToConsole "Completed"
Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTickfileManager Event Handlers
'@================================================================================

Private Sub mReplayController_NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "mReplayController_NotifyEvent"
On Error GoTo Err

LogMessage TypeName(ev.Source) & " :" & ev.EventMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mReplayController_ReplayProgress(ByVal pTickfileTimestamp As Date, ByVal pEventsPlayed As Long, ByVal pPercentComplete As Long)
Const ProcName As String = "mReplayController_ReplayProgress"
On Error GoTo Err

gCon.WriteLineToConsole "Processing tickfile: events processed=" & pEventsPlayed & "; percent complete=" & Format(pPercentComplete, "0.0")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub StartData( _
                ByVal pTB As TradeBuildAPI, _
                ByVal tickfileName As String, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal number As Long, _
                ByVal barLength As Long, _
                ByVal barUnits As TimePeriodUnits, _
                ByVal sessionOnly As Boolean)
Const ProcName As String = "startData"
On Error GoTo Err

Set mSelfRef = Me

Set mTB = pTB

mNumber = number
mBarLength = barLength
mBarUnits = barUnits
mSessionOnly = sessionOnly

Dim lTickfileSpec As New TickfileSpecifier
lTickfileSpec.FileName = tickfileName
lTickfileSpec.fromDate = fromDate
lTickfileSpec.toDate = toDate

Dim lTickfileSpecs As New TickFileSpecifiers
lTickfileSpecs.Add lTickfileSpec

Dim lTickDataManager As TickfileDataManager
Set lTickDataManager = CreateTickDataManager(lTickfileSpecs, _
                                            mTB.TickfileStoreInput, _
                                            mTB.ContractStorePrimary, _
                                            , _
                                            MarketDataSourceOptNoImpliedTrades + MarketDataSourceOptNoVolumeAdjustments, _
                                            , _
                                            , _
                                            0)

Dim lTickers As Tickers
Set lTickers = CreateTickers(lTickDataManager, mTB.StudyLibraryManager, mTB.HistoricalDataStoreInput)

Dim lticker As Ticker
Set lticker = lTickers.CreateTicker(mReplayController.TickStream(0).ContractFuture, False)

Dim lTimeframe As Timeframe
Set lTimeframe = lticker.Timeframes.AddHistorical(GetTimePeriod(mBarLength, mBarUnits), _
                                        "", _
                                        0, _
                                        0, _
                                        0, _
                                        , _
                                        Not mSessionOnly)
gProcessor.Timeframe = lTimeframe

Set mReplayController = lTickDataManager.ReplayController
Set mReplayTaskController = mReplayController.StartReplay

Set mTimer = CreateIntervalTimer(1000, , 1000)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub mTimer_TimerExpired(ev As TimerExpiredEventData)
Static lCounter As Long
lCounter = lCounter + 1
gCon.WriteLineToConsole "Timer tick " & lCounter
End Sub
