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

Private Const ModuleName                As String = "HistoricDataWriter"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type PendingBarWrite
    barType         As BarTypes
    barTimePeriod   As TimePeriod
    Bar             As Bar
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mWriter As TradeBuildSP.IBarDataWriter
Private mServiceProviderWriterReady As Boolean

Private mPendingWrites As Collection

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mSelfRef As HistoricDataWriter

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingWrites = New Collection
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

Friend Property Get ready() As Boolean
ready = mServiceProviderWriterReady
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ReleaseDataStore()
gLog pMsg:="Releasing data store", pProcName:="ReleaseDataStore", pProjName:=ProjectName, pModName:=ModuleName
mWriter.ReleaseDataStore

Set mSelfRef = Me
Set mTimer = CreateIntervalTimer(60, ExpiryTimeUnitMilliseconds)
mTimer.StartTimer
End Sub

Friend Sub ServiceProviderWriterNotReady()
mServiceProviderWriterReady = False
End Sub

Friend Sub ServiceProviderWriterReady()
Dim Item As Variant
Dim pendingWrite As PendingBarWrite

mServiceProviderWriterReady = True

For Each Item In mPendingWrites
    pendingWrite = Item
    WriteBar pendingWrite.Bar, _
            pendingWrite.barType, _
            pendingWrite.barTimePeriod
Next

Set mPendingWrites = New Collection
End Sub

Friend Sub WriteBar( _
                ByVal theBar As Bar, _
                ByVal barType As BarTypes, _
                ByVal barTimePeriod As TimePeriod)
Dim pendingWrite As PendingBarWrite

If Not mServiceProviderWriterReady Then
    pendingWrite.barType = barType
    Set pendingWrite.barTimePeriod = barTimePeriod
    Set pendingWrite.Bar = theBar
    mPendingWrites.AddToEnd pendingWrite
Else
    mWriter.WriteBar theBar, barType, barTimePeriod
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

