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

Private Const ProjectName                   As String = "gbd"
Private Const ModuleName                    As String = "FileBasedProcessor"

'@================================================================================
' Member variables
'@================================================================================

Private mSelfRef As Object

Private WithEvents mTickfileManager As TickFileManager
Attribute mTickfileManager.VB_VarHelpID = -1

Private mNumber As Long
Private mBarLength As Long
Private mBarUnits As TimePeriodUnits
Private mSessionOnly As Boolean

Private WithEvents mTimer As BaseIntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' mTickfileManager Event Handlers
'@================================================================================

Private Sub mTickfileManager_QueryReplayNextTickfile(ByVal tickfileIndex As Long, ByVal tickfileName As String, ByVal TickfileSizeBytes As Long, ByVal pContract As ContractUtils26.Contract, continueMode As TradeBuild26.ReplayContinueModes)
continueMode = ReplayContinue
End Sub

Private Sub mTickfileManager_ReplayCompleted()
gProcessor.Finish
Set gProcessor = Nothing
gCon.WriteLineToConsole "Completed"
Set mSelfRef = Nothing
End Sub

Private Sub mTickfileManager_ReplayProgress(ByVal tickfileTimestamp As Date, ByVal eventsPlayed As Long, ByVal percentComplete As Single)
gCon.WriteLineToConsole "Processing tickfile: events processed=" & eventsPlayed & "; percent complete=" & Format(percentComplete, "0.0")
End Sub

Private Sub mTickfileManager_TickerAllocated(ByVal pTicker As TradeBuild26.Ticker)
gProcessor.Ticker = pTicker

Dim lTimeframe As Timeframe
Set lTimeframe = pTicker.Timeframes.AddHistorical(GetTimePeriod(mBarLength, mBarUnits), _
                                        "", _
                                        0, _
                                        0, _
                                        0, _
                                        Not mSessionOnly)
gProcessor.Timeframe = lTimeframe
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub startData( _
                ByVal tickfileName As String, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal number As Long, _
                ByVal barLength As Long, _
                ByVal barUnits As TimePeriodUnits, _
                ByVal sessionOnly As Boolean)
                
Dim tfs As New TickfileSpecifier

Set mSelfRef = Me

mNumber = number
mBarLength = barLength
mBarUnits = barUnits
mSessionOnly = sessionOnly

tfs.FileName = tickfileName
tfs.fromDate = fromDate
tfs.toDate = toDate
Set mTickfileManager = TradeBuildAPI.Tickers.CreateTickFileManager(0)
mTickfileManager.TickfileSpecifier = tfs
mTickfileManager.ReplayProgressEventIntervalMillisecs = 1000
mTickfileManager.ReplaySpeed = 0
mTickfileManager.StartReplay

Set mTimer = CreateBaseIntervalTimer(1000, , 1000)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



Private Sub mTimer_TimerExpired()
Static lCounter As Long
lCounter = lCounter + 1
gCon.WriteLineToConsole "Timer tick " & lCounter
End Sub
