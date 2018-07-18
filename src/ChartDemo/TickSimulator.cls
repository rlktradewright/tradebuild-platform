VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements ITask

'================================================================================
' Events
'================================================================================

Event HistoricalBar( _
                ByVal timestamp As Date, _
                ByVal openPrice As Double, _
                ByVal highPrice As Double, _
                ByVal lowPrice As Double, _
                ByVal closePrice As Double, _
                ByVal volume As Long)
                
Event TickPrice( _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event TickVolume( _
                ByVal timestamp As Date, _
                ByVal volume As Long)
                

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "TickSimulator"

Private Const NumPriceChangeElements = 42
Private Const NumTickVolumeElements = 90

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mStartPrice As Double
Private mTickSize As Double

Private mPriceChange
Private mTickVolume

Private mTickCount As Long

Private mBarLength As Long

Private mCumVolume As Long

Private mTaskContext As TaskContext

Private mNumberOfBars As Long

Private mTimestamp As Date

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mPriceChange = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                    1, 1, 1, 1, 1, -1, -1, -1, -1, -1, _
                    2, 2, 2, -2, -2, -2, 3, 3, -3, -3, 4, -4)

mTickVolume = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
                2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
                3, 3, 3, 3, 3, 3, 3, 3, 3, 3, _
                4, 4, 4, 4, 4, 5, 5, 5, 10, 20)

mStartPrice = 1230
mTickSize = 0.25

End Sub

'================================================================================
' ITask Interface Members
'================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    mTaskContext.Finish Nothing, True
    Exit Sub
End If

Static sBarCount As Long

Dim bartime As Double
bartime = calcBarStartTime(mTimestamp)

Dim volume As Long
Dim closePrice As Double
Dim openPrice As Double
Dim highPrice As Double
Dim lowPrice As Double

lowPrice = 999999

Do
    closePrice = generateSimulatedPrice
    If openPrice = 0 Then openPrice = closePrice
    If closePrice > highPrice Then highPrice = closePrice
    If closePrice < lowPrice Then lowPrice = closePrice
    
    volume = volume + generateSimulatedTickVolume
    
    mTimestamp = mTimestamp + (Rnd * 5) / 86400
    If calcBarStartTime(mTimestamp) > bartime Then
        sBarCount = sBarCount + 1
        RaiseEvent HistoricalBar(bartime, _
                            openPrice, _
                            highPrice, _
                            lowPrice, _
                            closePrice, _
                            volume)
        mCumVolume = mCumVolume + volume
        Exit Do
    End If
Loop

If sBarCount >= mNumberOfBars Then
    Set mTimer = CreateIntervalTimer(2, ExpiryTimeUnitSeconds, 2000, True)
    mTimer.StartTimer
    mTaskContext.Finish Empty, False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mTimer_TimerExpired"
On Error GoTo Err

mTickCount = mTickCount + 1

Dim timestamp As Date
timestamp = Now

RaiseEvent TickPrice(timestamp, generateSimulatedPrice)

Dim volume As Long
volume = generateSimulatedTickVolume
mCumVolume = mCumVolume + volume
RaiseEvent TickVolume(timestamp, mCumVolume)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let BarLength(ByVal Value As Double)
Const ProcName As String = "BarLength"
On Error GoTo Err

mBarLength = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let StartPrice(ByVal Value As Double)
Const ProcName As String = "StartPrice"
On Error GoTo Err

mStartPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickCount() As Long
Const ProcName As String = "TickCount"
On Error GoTo Err

TickCount = mTickCount

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TickSize(ByVal Value As Double)
Const ProcName As String = "TickSize"
On Error GoTo Err

mTickSize = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Function StartSimulation(ByVal numberOfBars As Long) As TaskController
Const ProcName As String = "StartSimulation"
On Error GoTo Err

mNumberOfBars = numberOfBars
Randomize
mTimestamp = Now - (mNumberOfBars - 1) * mBarLength / 1440

Set StartSimulation = StartTask(Me, PriorityHigh)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub StopSimulation()
Const ProcName As String = "StopSimulation"
On Error GoTo Err

If Not mTimer Is Nothing Then mTimer.StopTimer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Function calcBarStartTime(ByVal timestamp As Date) As Date
Const ProcName As String = "calcBarStartTime"
On Error GoTo Err

calcBarStartTime = Int(CDbl(timestamp) * 1440 / mBarLength) * mBarLength / 1440

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateSimulatedPrice() As Double
Const ProcName As String = "generateSimulatedPrice"
On Error GoTo Err

Static sPrevprice As Double
If sPrevprice = 0 Then sPrevprice = mStartPrice
Randomize
generateSimulatedPrice = sPrevprice + mTickSize * mPriceChange(Fix(Rnd() * NumPriceChangeElements))
sPrevprice = generateSimulatedPrice

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateSimulatedTickVolume() As Long
Const ProcName As String = "generateSimulatedTickVolume"
On Error GoTo Err

Randomize
generateSimulatedTickVolume = mTickVolume(Fix(Rnd() * NumTickVolumeElements))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


