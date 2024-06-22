VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarDataProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "BarDataProcessor"

Private Const ColumnNameDateTime                    As String = "DateTime"
Private Const ColumnNameOpenValue                   As String = "OpenPrice"
Private Const ColumnNameHighValue                   As String = "HighPrice"
Private Const ColumnNameLowValue                    As String = "LowPrice"
Private Const ColumnNameCloseValue                  As String = "ClosePrice"
Private Const ColumnNameVolume                      As String = "Volume"
Private Const ColumnNameTickVolume                  As String = "TickVolume"

'@================================================================================
' Member variables
'@================================================================================

Private mRS                                         As Recordset

Private mSessionTimes                               As SessionTimes

Private mStartAtFromDate                            As Boolean

Private mMaxNumberOfBars                            As Long

Private mAccumulatedVolume                          As Long

Private mTargetBarsBuilder                          As BarsBuilder

Private mSessionBuilder                             As SessionBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTargetBarsBuilder As BarsBuilder, _
                ByVal pRS As Recordset, _
                ByRef pSessionTimes As SessionTimes, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pStartAtFromDate As Boolean)
Set mTargetBarsBuilder = pTargetBarsBuilder
Set mRS = pRS
mSessionTimes = pSessionTimes
Set mSessionBuilder = CreateSessionBuilder(mSessionTimes.StartTime, mSessionTimes.EndTime)
mMaxNumberOfBars = pMaxNumberOfBars
mStartAtFromDate = pStartAtFromDate
End Sub

Public Function ProcessNextBar() As Boolean
Const ProcName As String = "ProcessNextBar"
On Error GoTo Err

If mRS.EOF Then Exit Function

ProcessBar mRS(ColumnNameDateTime), mRS(ColumnNameOpenValue), mRS(ColumnNameHighValue), mRS(ColumnNameLowValue), mRS(ColumnNameCloseValue), mRS(ColumnNameVolume), mRS(ColumnNameTickVolume)

mRS.MoveNext
ProcessNextBar = True

Exit Function

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ProcessBars()
Const ProcName As String = "ProcessBars"
On Error GoTo Err

Do While ProcessNextBar
Loop

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub ProcessBar( _
    ByVal pTimestamp As Date, _
    ByVal pOpenValue As Double, _
    ByVal pHighValue As Double, _
    ByVal pLowValue As Double, _
    ByVal pCloseValue As Double, _
    ByVal pVolume As Long, _
    ByVal pTickVolume As Long)
Const ProcName As String = "ProcessBar"
On Error GoTo Err

Static sCurrentBarEndTime As Date
Static sCurrentBarNumber As Long

setSessionTime pTimestamp

' ignore bars containing spurious large volumes collected from
' TWS for some futures during June - October 2021
If pVolume > 1000000000 Then Exit Sub

Dim lBarEndTime As Date
lBarEndTime = BarEndTime(pTimestamp, mTargetBarsBuilder.BarTimePeriod, mSessionTimes.StartTime, mSessionTimes.EndTime)
If lBarEndTime > sCurrentBarEndTime Then
    sCurrentBarEndTime = lBarEndTime
    sCurrentBarNumber = sCurrentBarNumber + 1
    If mStartAtFromDate And sCurrentBarNumber > mMaxNumberOfBars Then Exit Sub
End If

mTargetBarsBuilder.NotifyBar pTimestamp, _
                            pOpenValue, _
                            pHighValue, _
                            pLowValue, _
                            pCloseValue, _
                            pVolume, _
                            pTickVolume
                            
mAccumulatedVolume = mAccumulatedVolume + pVolume

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSessionTime( _
                ByVal timestamp As Date)
Const ProcName As String = "setSessionTime"
On Error GoTo Err

Dim lSessEv As SessionEventData
lSessEv = mSessionBuilder.SetSessionCurrentTime(timestamp)
If lSessEv.ChangeType = SessionChangeStart Then
    mAccumulatedVolume = 0
End If

Exit Sub

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Sub




