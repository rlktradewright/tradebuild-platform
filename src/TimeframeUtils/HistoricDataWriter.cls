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
    BarTimePeriod   As TimePeriod
    Bar             As Bar
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mWriter As IBarDataWriter
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

Private Sub mTimer_TimerExpired(ev As TimerExpiredEventData)
Set mWriter = Nothing
Set mSelfRef = Nothing
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ServiceProviderWriter(ByVal Value As IBarDataWriter)
Set mWriter = Value
End Property

Friend Property Get ready() As Boolean
ready = mServiceProviderWriterReady
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ReleaseDataStore()
Const ProcName As String = "ReleaseDataStore"
On Error GoTo Err

gLogLogger.Log pMsg:="Releasing Data store", pProcName:=ProcName, pModName:=ModuleName
mWriter.ReleaseDataStore

Set mSelfRef = Me
Set mTimer = CreateIntervalTimer(60, ExpiryTimeUnitMilliseconds)
mTimer.StartTimer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ServiceProviderWriterNotReady()
Const ProcName As String = "ServiceProviderWriterNotReady"
On Error GoTo Err

mServiceProviderWriterReady = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ServiceProviderWriterReady()
Dim Item As Variant
Dim pendingWrite As PendingBarWrite

Const ProcName As String = "ServiceProviderWriterReady"
On Error GoTo Err

mServiceProviderWriterReady = True

For Each Item In mPendingWrites
    pendingWrite = Item
    WriteBar pendingWrite.Bar, _
            pendingWrite.barType, _
            pendingWrite.BarTimePeriod
Next

Set mPendingWrites = New Collection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub WriteBar( _
                ByVal theBar As Bar, _
                ByVal barType As BarTypes, _
                ByVal BarTimePeriod As TimePeriod)
Dim pendingWrite As PendingBarWrite

Const ProcName As String = "WriteBar"
On Error GoTo Err

If Not mServiceProviderWriterReady Then
    pendingWrite.barType = barType
    Set pendingWrite.BarTimePeriod = BarTimePeriod
    Set pendingWrite.Bar = theBar
    mPendingWrites.AddToEnd pendingWrite
Else
    mWriter.WriteBar theBar, barType, BarTimePeriod
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

