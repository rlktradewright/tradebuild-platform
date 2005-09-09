VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoricDataWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mWriter As TradeBuildSP.IBarDataWriter
Private mServiceProviderWriterReady As Boolean

Private mPendingWrites As Collection

Private WithEvents mTimer As TimerUtils.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mSelfRef As TradeBuild.HistoricDataWriter

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mPendingWrites = New Collection
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Set mWriter = Nothing
Set mSelfRef = Nothing
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let ServiceProviderWriter(ByVal value As TradeBuildSP.IBarDataWriter)
Set mWriter = value
End Property

Friend Property Let ServiceProviderWriterReady(ByVal value As Boolean)
Dim pendingWrite As Bar

mServiceProviderWriterReady = value

If mPendingWrites.Count <> 0 Then
    For Each pendingWrite In mPendingWrites
        WriteBar pendingWrite
    Next
End If

Set mPendingWrites = Nothing
Set mPendingWrites = New Collection
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub ReleaseDataStore()
mWriter.ReleaseDataStore

Set mSelfRef = Me
Set mTimer = New TimerUtils.IntervalTimer
mTimer.RepeatNotifications = False
mTimer.TimerIntervalSecs = 60
mTimer.StartTimer
End Sub

Friend Sub WriteBar(ByVal theBar As TradeBuild.Bar)
Dim spBar As TradeBuildSP.Bar

If Not mServiceProviderWriterReady Then
    mPendingWrites.add theBar
Else
    With spBar
        .barType = theBar.barType
        .closePrice = theBar.closeValue
        .highPrice = theBar.highValue
        .lowPrice = theBar.lowValue
        .openPrice = theBar.openValue
        .PeriodMinutes = theBar.barLength
        .tickVolume = theBar.tickVolume
        .Timestamp = theBar.datetime
        .volume = theBar.volume
    End With
    mWriter.WriteBar spBar
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

