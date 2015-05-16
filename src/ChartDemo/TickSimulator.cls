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
' XXXX Interface Members
'================================================================================

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired(ev As TimerExpiredEventData)
Dim timestamp As Date
Dim volume As Long

Const ProcName As String = "mTimer_TimerExpired"
On Error GoTo Err

mTickCount = mTickCount + 1
timestamp = Now
RaiseEvent TickPrice(timestamp, generateSimulatedPrice)
volume = generateSimulatedTickVolume
mCumVolume = mCumVolume + volume
RaiseEvent TickVolume(timestamp, mCumVolume)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let BarLength(ByVal value As Double)
Const ProcName As String = "BarLength"
On Error GoTo Err

mBarLength = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let StartPrice(ByVal value As Double)
Const ProcName As String = "StartPrice"
On Error GoTo Err

mStartPrice = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TickCount() As Long
Const ProcName As String = "TickCount"
On Error GoTo Err

TickCount = mTickCount

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TickSize(ByVal value As Double)
Const ProcName As String = "TickSize"
On Error GoTo Err

mTickSize = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub StartSimulation(ByVal numberOfBars As Long)
Dim barcount As Long
Dim timestamp As Date
Dim bartime As Double
Dim openPrice As Double
Dim highPrice As Double
Dim lowPrice As Double
Dim closePrice As Double
Dim volume As Long

Const ProcName As String = "StartSimulation"
On Error GoTo Err

Randomize

timestamp = Now - (numberOfBars - 1) * mBarLength / 1440
bartime = calcBarTime(timestamp)
lowPrice = 999999

Do While barcount < numberOfBars
    closePrice = generateSimulatedPrice
    If openPrice = 0 Then openPrice = closePrice
    If closePrice > highPrice Then highPrice = closePrice
    If closePrice < lowPrice Then lowPrice = closePrice
    
    volume = volume + generateSimulatedTickVolume
    
    timestamp = timestamp + (Rnd * 5) / 86400
    If calcBarTime(timestamp) > bartime Then
        barcount = barcount + 1
        RaiseEvent HistoricalBar(bartime, _
                            openPrice, _
                            highPrice, _
                            lowPrice, _
                            closePrice, _
                            volume)
        bartime = calcBarTime(timestamp)
        mCumVolume = mCumVolume + volume
        openPrice = 0
        highPrice = 0
        lowPrice = 999999
        closePrice = 0
        volume = 0
        
        If barcount Mod 50 = 0 Then
            DoEvents
        End If
    End If
Loop

Set mTimer = CreateIntervalTimer(2, ExpiryTimeUnitSeconds, 2000, True)
mTimer.StartTimer

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Public Sub StopSimulation()
Const ProcName As String = "StopSimulation"
On Error GoTo Err

If Not mTimer Is Nothing Then mTimer.StopTimer

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Function calcBarTime(ByVal timestamp As Date) As Date
Const ProcName As String = "calcBarTime"
On Error GoTo Err

calcBarTime = Int(CDbl(timestamp) * 1440 / mBarLength) * mBarLength / 1440

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function generateSimulatedPrice() As Double
Static prevprice As Double
Const ProcName As String = "generateSimulatedPrice"
On Error GoTo Err

If prevprice = 0 Then prevprice = mStartPrice
Randomize
generateSimulatedPrice = prevprice + mTickSize * mPriceChange(Fix(Rnd() * NumPriceChangeElements))
prevprice = generateSimulatedPrice

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function generateSimulatedTickVolume() As Long
Const ProcName As String = "generateSimulatedTickVolume"
On Error GoTo Err

Randomize
generateSimulatedTickVolume = mTickVolume(Fix(Rnd() * NumTickVolumeElements))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function


