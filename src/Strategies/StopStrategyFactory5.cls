VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopStrategyFactory5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IPosnMgmtStrtgyFactory

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

Private Const ModuleName                            As String = "StopStrategyFactory5"

'@================================================================================
' Member variables
'@================================================================================

Private S                                           As New StopStrategy5Static

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IPosnMgmtStrtgyFactory Interface Members
'@================================================================================

Private Function IPosnMgmtStrtgyFactory_CreateStrategy( _
                ByVal pTradingContext As TradingContext) As IPositionMgmtStrategy
Const ProcName As String = "IPosnMgmtStrtgyFactory_CreateStrategy"
On Error GoTo Err

Dim l As New StopStrategy5
l.Initialise S
Set IPosnMgmtStrtgyFactory_CreateStrategy = l

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IPosnMgmtStrtgyFactory_DefineDefaultParameters()
Const ProcName As String = "IPosnMgmtStrtgyFactory_DefineDefaultParameters"
On Error GoTo Err

DefineParameter ParamBarLength, BarLength
DefineParameter ParamBarUnit, BarUnit
DefineParameter ParamIncludeBarsOutsideSession, IncludeBarsOutsideSession
DefineParameter ParamBollingerPeriods, BollingerPeriods
DefineParameter ParamBollingerMovingAverageType, "SMA"
DefineParameter ParamMinimumSwingTicks, MinimumSwingTicks
DefineParameter ParamRetraceFromExtremes, RetraceFromExtremes
DefineParameter ParamBreakoutThresholdTicks, StopBreakoutThresholdTicks
DefineParameter ParamLongBarLengthFactor, LongBarLengthFactor
DefineParameter ParamLongBollPeriods, LongBollPeriods
DefineParameter ParamLongMAPeriods, LongMAPeriods
DefineParameter ParamBackstopBarLengthFactor, BackstopBarLengthFactor
DefineParameter ParamBackstopMAPeriods, BackstopMAPeriods
DefineParameter ParamBreakEvenBarLength, BreakevenBarLength
DefineParameter ParamRetracementStopPercent, RetracementStopPercent
DefineParameter ParamRetracementStopThresholdTicks, RetracementStopThresholdTicks
DefineParameter ParamBreakEvenBarThresholdTicks, BreakevenBarThresholdTicks
DefineParameter ParamBreakEvenThresholdTicks, BreakEvenThresholdTicks
DefineParameter ParamConfirmedBarsForBreakEven, ConfirmedBarsForBreakEven
DefineParameter ParamConfirmedBarsForSwing, ConfirmedBarsForSwing
DefineParameter ParamConfirmedBarsForTrail, ConfirmedBarsForTrail
DefineParameter ParamConfirmedBarsForExit, ConfirmedBarsForExit
DefineParameter ParamBollingerCentreBandWidthTicks, BollCentreBandwidthTicks
DefineParameter ParamBollingerEdgeBandWidthTicks, BollEdgeBandwidthTicks
DefineParameter ParamSwingToMoveStopTicks, SwingToMoveStopTicks
DefineParameter ParamMaxContraSwingPercent, MaxContraSwingPercent
DefineParameter ParamTicksSwingToTrail, TicksSwingToTrail
DefineParameter ParamUseIntermediateStops, UseIntermediateStops
DefineParameter ParamRewardToRiskRatio, RewardToRiskRatio

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IPosnMgmtStrtgyFactory_Initialise( _
                ByVal pContext As InitialisationContext)
Const ProcName As String = "IPosnMgmtStrtgyFactory_Initialise"
On Error GoTo Err

processParameters

Dim lBarLength As Long
lBarLength = GetParameterValue(ParamBarLength)

Dim lBarUnit As String
lBarUnit = GetParameterValue(ParamBarUnit)

Dim lBaseTimeframe As ResourceIdentifier
Set lBaseTimeframe = DeclareTimeframe(lBarLength, _
                                lBarUnit, _
                                500, _
                                GetParameterValue(ParamIncludeBarsOutsideSession))

Dim lMinimumSwingTicks As Long
lMinimumSwingTicks = CLng(GetParameterValue(ParamMinimumSwingTicks, MinimumSwingTicks))
If lMinimumSwingTicks <> 0 Then
    DeclareStudy "Swing", _
                lBaseTimeframe, _
                "Minimum swing (ticks)=" & lMinimumSwingTicks & _
                    ";Include implicit swing points=Y", _
                GetParameterValue(ParamIncludeBarsOutsideSession), _
                3
End If

Dim lLongBarLengthFactor As Long
lLongBarLengthFactor = CLng(GetParameterValue(ParamLongBarLengthFactor, LongBarLengthFactor))

Dim lLongBollPeriods As Long
lLongBollPeriods = CLng(GetParameterValue(ParamLongBollPeriods, LongBollPeriods))

Dim lLongMAPeriods As Long
lLongMAPeriods = CLng(GetParameterValue(ParamLongMAPeriods, LongMAPeriods))

If lLongBarLengthFactor <> 0 And _
    (lLongMAPeriods <> 0 Or lLongBollPeriods <> 0) _
Then
    Set S.LongTimeframe = DeclareTimeframe(lLongBarLengthFactor * lBarLength, _
                                    lBarUnit, _
                                    IIf(lLongMAPeriods > lLongBollPeriods, lLongMAPeriods, lLongBollPeriods) + 1)
    
    If lLongMAPeriods <> 0 Then
        Set S.LongMA = DeclareStudy("Simple Moving Average", _
                                S.LongTimeframe, _
                                "Periods=" & lLongMAPeriods, _
                                GetParameterValue(ParamIncludeBarsOutsideSession), _
                                3)
    End If
    
    If lLongBollPeriods <> 0 Then
        Set S.LongBoll = DeclareStudy("Bollinger Bands", _
                                S.LongTimeframe, _
                                "Periods=" & lLongBollPeriods & _
                                    ";Standard deviations=2", _
                                GetParameterValue(ParamIncludeBarsOutsideSession), _
                                3)
    End If
End If

If S.ConfirmedBarsForBreakEven <> 0 Or _
    S.ConfirmedBarsForSwing <> 0 Or _
    S.ConfirmedBarsForTrailing <> 0 Or _
    S.ConfirmedBarsForExit <> 0 _
Then
    DeclareStudy "MACD", _
                lBaseTimeframe, _
                "Short periods=5" & _
                    ";Long periods=34" & _
                    ";Mov avg type=SMA" & _
                    ";Smoothing periods=5", _
                GetParameterValue(ParamIncludeBarsOutsideSession), _
                10
End If

Dim lBackstopBarLengthFactor As Long
lBackstopBarLengthFactor = GetParameterValue(ParamBackstopBarLengthFactor)

Dim lBackstopMAPeriods As Long
lBackstopMAPeriods = GetParameterValue(ParamBackstopMAPeriods)

If lBackstopBarLengthFactor > 0 And _
    lBackstopMAPeriods > 0 _
Then
    Set S.StopTimeframe = DeclareTimeframe(lBackstopBarLengthFactor * lBarLength, _
                                                lBarUnit, _
                                                lBackstopMAPeriods + 1)

    Set S.StopMA = DeclareStudy("Exponential Moving Average", _
                            S.StopTimeframe, _
                            "Periods=" & lBackstopMAPeriods, _
                            GetParameterValue(ParamIncludeBarsOutsideSession), _
                            2, _
                            InputNames:="HL/2")

    Set S.StopMA = DeclareStudy("Bollinger Bands", _
                            S.StopTimeframe, _
                            "Periods=34" & _
                                ";Standard deviations=2", _
                            GetParameterValue(ParamIncludeBarsOutsideSession), _
                            2)
End If
    
If GetParameterValue(ParamBreakEvenBarLength) > 0 Then
    Set S.BreakevenTimeframe = DeclareTimeframe(GetParameterValue(ParamBreakEvenBarLength), _
                                                GetParameterValue(ParamBarUnit), _
                                                2)
End If




Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processParameters()
Const ProcName As String = "processParameters"
On Error GoTo Err

S.BreakEvenThresholdTicks = GetParameterValue(ParamBreakEvenThresholdTicks)
S.BreakevenBarThresholdTicks = GetParameterValue(ParamBreakEvenBarThresholdTicks)
S.BreakoutThresholdTicks = GetParameterValue(ParamBreakoutThresholdTicks)
S.ConfirmedBarsForBreakEven = GetParameterValue(ParamConfirmedBarsForBreakEven)
S.ConfirmedBarsForSwing = GetParameterValue(ParamConfirmedBarsForSwing)
S.ConfirmedBarsForTrailing = GetParameterValue(ParamConfirmedBarsForTrail)
S.ConfirmedBarsForExit = GetParameterValue(ParamConfirmedBarsForExit)
S.RetraceFromExtremes = CBool(GetParameterValue(ParamRetraceFromExtremes))

Dim v As Double
v = CDbl(GetParameterValue(ParamMaxContraSwingPercent, "0"))
If v = 0 Then
    S.MaxContraSwingFactor = 0
Else
    S.MaxContraSwingFactor = 1! - v / 100
End If

v = CDbl(GetParameterValue(ParamRetracementStopPercent, "0"))
If v = 0# Then
    S.RetracementStopFactor = 0#
Else
    S.RetracementStopFactor = 1# - v / 100#
End If

S.RetracementStopThresholdTicks = GetParameterValue(ParamRetracementStopThresholdTicks)
S.RewardToRiskRatio = GetParameterValue(ParamRewardToRiskRatio)
S.SwingToMoveStopTicks = GetParameterValue(ParamSwingToMoveStopTicks)
S.TicksSwingToTrail = GetParameterValue(ParamTicksSwingToTrail)
S.UseIntermediateStops = GetParameterValue(ParamUseIntermediateStops)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




