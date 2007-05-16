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
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type PendingBarWrite
    barType         As TradeBuildSP.BarTypes
    barLength       As Long
    barUnit         As TimePeriodUnits
    Bar             As StudyUtils26.Bar
    volume          As Long
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mWriter As TradeBuildSP.IBarDataWriter
Private mServiceProviderWriterReady As Boolean

Private mPendingWrites As Queue

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mSelfRef As TradeBuild26.HistoricDataWriter

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingWrites = New Queue
mPendingWrites.initialCapacity = 20
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mTimer Event Handlers
'@================================================================================

Private Sub mTimer_TimerExpired()
Set mWriter = Nothing
Set mSelfRef = Nothing
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ServiceProviderWriter(ByVal value As TradeBuildSP.IBarDataWriter)
Set mWriter = value
End Property

Friend Property Let ServiceProviderWriterReady(ByVal value As Boolean)
Dim pendingWrite As PendingBarWrite

mServiceProviderWriterReady = value

If Not mServiceProviderWriterReady Then Exit Property
    
Do While mPendingWrites.size <> 0
    pendingWrite = mPendingWrites.removeFromFront
    WriteBar pendingWrite.Bar, _
            pendingWrite.barType, _
            pendingWrite.volume, _
            pendingWrite.barLength, _
            pendingWrite.barUnit
Loop

End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ReleaseDataStore()
mWriter.ReleaseDataStore

Set mSelfRef = Me
Set mTimer = createIntervalTimer(60, ExpiryTimeUnitMilliseconds)
mTimer.StartTimer
End Sub

Friend Sub WriteBar( _
                ByVal theBar As StudyUtils26.Bar, _
                ByVal barType As TradeBuildSP.BarTypes, _
                ByVal volume As Long, _
                ByVal barLength As Long, _
                ByVal barUnit As TimePeriodUnits)
Dim spBar As TradeBuildSP.Bar
Dim pendingWrite As PendingBarWrite

If Not mServiceProviderWriterReady Then
    pendingWrite.barType = barType
    pendingWrite.barLength = barLength
    pendingWrite.barUnit = barUnit
    Set pendingWrite.Bar = theBar
    pendingWrite.volume = volume
    mPendingWrites.addToEnd pendingWrite
Else
    With spBar
        .closePrice = theBar.closeValue
        .highPrice = theBar.highValue
        .lowPrice = theBar.lowValue
        .OpenPrice = theBar.openValue
        .tickVolume = theBar.tickVolume
        .timestamp = theBar.DateTime
        .volume = volume
    End With
    mWriter.WriteBar spBar, barType, barLength, barUnit
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

