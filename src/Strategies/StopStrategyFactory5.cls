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

Private mParams                                     As Parameters

Private S                                           As New StopStrategy5Static

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStopStrategyFactory Interface Members
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

Private Property Get IPosnMgmtStrtgyFactory_DefaultParameters() As Parameters
Const ProcName As String = "IPosnMgmtStrtgyFactory_DefaultParameters"
On Error GoTo Err

Dim lParams As New Parameters
lParams.SetParameterValue ParamBarLength, BarLength
lParams.SetParameterValue ParamBarUnit, BarUnit
lParams.SetParameterValue ParamIncludeBarsOutsideSession, IncludeBarsOutsideSession
lParams.SetParameterValue ParamBollingerPeriods, BollingerPeriods
lParams.SetParameterValue ParamBollingerMovingAverageType, "SMA"
lParams.SetParameterValue ParamMinimumSwingTicks, MinimumSwingTicks
lParams.SetParameterValue ParamRetraceFromExtremes, RetraceFromExtremes
lParams.SetParameterValue ParamBreakoutThresholdTicks, StopBreakoutThresholdTicks
lParams.SetParameterValue ParamLongBarLengthFactor, LongBarLengthFactor
lParams.SetParameterValue ParamLongBollPeriods, LongBollPeriods
lParams.SetParameterValue ParamLongMAPeriods, LongMAPeriods
lParams.SetParameterValue ParamBackstopBarLengthFactor, BackstopBarLengthFactor
lParams.SetParameterValue ParamBackstopMAPeriods, BackstopMAPeriods
lParams.SetParameterValue ParamBreakEvenBarLength, BreakevenBarLength
lParams.SetParameterValue ParamRetracementStopPercent, RetracementStopPercent
lParams.SetParameterValue ParamRetracementStopThresholdTicks, RetracementStopThresholdTicks
lParams.SetParameterValue ParamBreakEvenBarThresholdTicks, BreakevenBarThresholdTicks
lParams.SetParameterValue ParamBreakEvenThresholdTicks, BreakEvenThresholdTicks
lParams.SetParameterValue ParamConfirmedBarsForBreakEven, ConfirmedBarsForBreakEven
lParams.SetParameterValue ParamConfirmedBarsForSwing, ConfirmedBarsForSwing
lParams.SetParameterValue ParamConfirmedBarsForTrail, ConfirmedBarsForTrail
lParams.SetParameterValue ParamConfirmedBarsForExit, ConfirmedBarsForExit
lParams.SetParameterValue ParamBollingerCentreBandWidthTicks, BollCentreBandwidthTicks
lParams.SetParameterValue ParamBollingerEdgeBandWidthTicks, BollEdgeBandwidthTicks
lParams.SetParameterValue ParamSwingToMoveStopTicks, SwingToMoveStopTicks
lParams.SetParameterValue ParamMaxContraSwingPercent, MaxContraSwingPercent
lParams.SetParameterValue ParamTicksSwingToTrail, TicksSwingToTrail
lParams.SetParameterValue ParamUseIntermediateStops, UseIntermediateStops
lParams.SetParameterValue ParamRewardToRiskRatio, RewardToRiskRatio

Set IPosnMgmtStrtgyFactory_DefaultParameters = lParams

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IPosnMgmtStrtgyFactory_Initialise( _
                ByVal pContext As InitialisationContext, _
                ByVal pParams As Parameters)
Const ProcName As String = "IPosnMgmtStrtgyFactory_Initialise"
On Error GoTo Err

processParameters pParams

Dim lBarLength As Long
lBarLength = mParams.GetParameterValue(ParamBarLength)

Dim lBarUnit As TimePeriodUnits
lBarUnit = TimePeriodUnitsFromString(mParams.GetParameterValue(ParamBarUnit))
Set S.Timeframe = pContext.AddTimeframe(GetTimePeriod(lBarLength, _
                                                lBarUnit), _
                                500, _
                                mParams.GetParameterValue(ParamIncludeBarsOutsideSession))

ReDim linputvaluenames(0) As String

Dim params As Parameters

Dim lMinimumSwingTicks As Long
lMinimumSwingTicks = CLng(mParams.GetParameterValue(ParamMinimumSwingTicks, MinimumSwingTicks))
If lMinimumSwingTicks <> 0 Then
    Set params = New Parameters
    params.SetParameterValue "Minimum swing (ticks)", lMinimumSwingTicks
    params.SetParameterValue "Include implicit swing points", "Y"
    linputvaluenames(0) = "$default"
    Set S.Swing = pContext.AddStudy("Swing", _
                                S.Timeframe.BarStudy, _
                                linputvaluenames, _
                                mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                                params, _
                                3)
End If

Dim lLongBarLengthFactor As Long
lLongBarLengthFactor = CLng(mParams.GetParameterValue(ParamLongBarLengthFactor, LongBarLengthFactor))

Dim lLongBollPeriods As Long
lLongBollPeriods = CLng(mParams.GetParameterValue(ParamLongBollPeriods, LongBollPeriods))

Dim lLongMAPeriods As Long
lLongMAPeriods = CLng(mParams.GetParameterValue(ParamLongMAPeriods, LongMAPeriods))

If lLongBarLengthFactor <> 0 And _
    (lLongMAPeriods <> 0 Or lLongBollPeriods <> 0) _
Then
    Set S.LongTimeframe = pContext.AddTimeframe(GetTimePeriod(lLongBarLengthFactor * lBarLength, _
                                                            lBarUnit), _
                                    IIf(lLongMAPeriods > lLongBollPeriods, lLongMAPeriods, lLongBollPeriods) + 1)
    
    If lLongMAPeriods <> 0 Then
        Set params = New Parameters
        params.SetParameterValue "Periods", lLongMAPeriods
        linputvaluenames(0) = "$default"
        Set S.LongMA = pContext.AddStudy("Simple Moving Average", _
                                S.LongTimeframe.BarStudy, _
                                linputvaluenames, _
                                mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                                params, _
                                3)
    End If
    
    If lLongBollPeriods <> 0 Then
        Set params = New Parameters
        params.SetParameterValue "Periods", lLongBollPeriods
        params.SetParameterValue "Standard deviations", 2
        Set S.LongBoll = pContext.AddStudy("Bollinger Bands", _
                                S.LongTimeframe.BarStudy, _
                                linputvaluenames, _
                                mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                                params, _
                                3)
    End If
End If

If S.ConfirmedBarsForBreakEven <> 0 Or _
    S.ConfirmedBarsForSwing <> 0 Or _
    S.ConfirmedBarsForTrailing <> 0 Or _
    S.ConfirmedBarsForExit <> 0 _
Then
    Set params = New Parameters
    params.SetParameterValue "Short periods", 5
    params.SetParameterValue "Long periods", 34
    params.SetParameterValue "Mov avg type", "SMA"
    params.SetParameterValue "Smoothing periods", 5
    linputvaluenames(0) = "$default"
    Set S.MACDStudy = pContext.AddStudy("MACD", _
                            S.Timeframe.BarStudy, _
                            linputvaluenames, _
                            mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                            params, _
                            10)
End If

Dim lBackstopBarLengthFactor As Long
lBackstopBarLengthFactor = mParams.GetParameterValue(ParamBackstopBarLengthFactor)

Dim lBackstopMAPeriods As Long
lBackstopMAPeriods = mParams.GetParameterValue(ParamBackstopMAPeriods)

If lBackstopBarLengthFactor > 0 And _
    lBackstopMAPeriods > 0 _
Then
    Set S.StopTimeframe = pContext.AddTimeframe(GetTimePeriod(lBackstopBarLengthFactor * lBarLength, _
                                                            lBarUnit), _
                                                lBackstopMAPeriods + 1)

    Set params = New Parameters
    params.SetParameterValue "Periods", lBackstopMAPeriods
    linputvaluenames(0) = "HL/2"
    Set S.StopMA = pContext.AddStudy("Exponential Moving Average", _
                            S.StopTimeframe.BarStudy, _
                            linputvaluenames, _
                            mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                            params, _
                            2)

    Set params = New Parameters
    params.SetParameterValue "Periods", 34
    params.SetParameterValue "Standard deviations", 2
    linputvaluenames(0) = "$default"
    Set S.StopMA = pContext.AddStudy("Bollinger Bands", _
                            S.StopTimeframe.BarStudy, _
                            linputvaluenames, _
                            mParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                            params, _
                            2)
End If
    
If mParams.GetParameterValue(ParamBreakEvenBarLength) > 0 Then
    Set S.BreakevenTimeframe = pContext.AddTimeframe(GetTimePeriod(mParams.GetParameterValue(ParamBreakEvenBarLength), _
                                                                mParams.GetParameterValue(ParamBarUnit)), _
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

Private Sub processParameters(ByVal pParams As Parameters)
Const ProcName As String = "processParameters"
On Error GoTo Err

Set mParams = pParams

S.BreakEvenThresholdTicks = mParams.GetParameterValue(ParamBreakEvenThresholdTicks)
S.BreakevenBarThresholdTicks = mParams.GetParameterValue(ParamBreakEvenBarThresholdTicks)
S.BreakoutThresholdTicks = mParams.GetParameterValue(ParamBreakoutThresholdTicks)
S.ConfirmedBarsForBreakEven = mParams.GetParameterValue(ParamConfirmedBarsForBreakEven)
S.ConfirmedBarsForSwing = mParams.GetParameterValue(ParamConfirmedBarsForSwing)
S.ConfirmedBarsForTrailing = mParams.GetParameterValue(ParamConfirmedBarsForTrail)
S.ConfirmedBarsForExit = mParams.GetParameterValue(ParamConfirmedBarsForExit)
S.RetraceFromExtremes = CBool(mParams.GetParameterValue(ParamRetraceFromExtremes))

Dim v As Double
v = CDbl(mParams.GetParameterValue(ParamMaxContraSwingPercent, "0"))
If v = 0 Then
    S.MaxContraSwingFactor = 0
Else
    S.MaxContraSwingFactor = 1! - v / 100
End If

v = CDbl(mParams.GetParameterValue(ParamRetracementStopPercent, "0"))
If v = 0# Then
    S.RetracementStopFactor = 0#
Else
    S.RetracementStopFactor = 1# - v / 100#
End If

S.RetracementStopThresholdTicks = mParams.GetParameterValue(ParamRetracementStopThresholdTicks)
S.RewardToRiskRatio = mParams.GetParameterValue(ParamRewardToRiskRatio)
S.SwingToMoveStopTicks = mParams.GetParameterValue(ParamSwingToMoveStopTicks)
S.TicksSwingToTrail = mParams.GetParameterValue(ParamTicksSwingToTrail)
S.UseIntermediateStops = mParams.GetParameterValue(ParamUseIntermediateStops)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




