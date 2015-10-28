Attribute VB_Name = "Globals"
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

Public Const ProjectName                            As String = "Strategies27"
Private Const ModuleName                            As String = "Globals"

Public Const ATRPeriods As Long = 20

Public Const BackstopBarLengthFactor As Long = 0
Public Const BackstopMAPeriods As Long = 13
Public Const BarLength As Long = 1
Public Const BarUnit As String = "min"
Public Const BollCentreBandwidthTicks As Single = 20
Public Const BollEdgeBandwidthTicks As Single = 10
Public Const BollingerMovingAverageType As String = "SMA"
Public Const BollingerPeriods As Long = 34
Public Const BreakevenBarLength As Long = 0
Public Const BreakevenBarThresholdTicks As Long = 0
Public Const BreakEvenThresholdTicks As Long = 0

Public Const ConfirmedBarsForBreakEven As Long = 0
Public Const ConfirmedBarsForExit As Long = 0
Public Const ConfirmedBarsForSwing As Long = 0
Public Const ConfirmedBarsForTrail As Long = 0

Public Const EntryBreakoutThresholdTicks = 4
Public Const EntryLimitOffsetTicks As Integer = -1

Public Const IncludeBarsOutsideSession As Boolean = False
Public Const InitialStopFactor As Double = 2#

Public Const LongBarLengthFactor As Long = 0
Public Const LongBollPeriods As Long = 0
Public Const LongMAPeriods As Long = 0

Public Const MaxContraSwingPercent As Double = 0#
Public Const MaxIncrements As Long = 3

' max number of ticks from entry price for initial stop loss
Public Const MaxInitialStopTicks = 100

Public Const MaxTradeSize As Long = 1
Public Const MinimumSwingTicks As Integer = 10

Public Const RetraceFromExtremes As Boolean = True
Public Const RetracementStopPercent As Single = 0
Public Const RetracementStopThresholdTicks As Long = 0
Public Const RewardToRiskRatio As Double = 0#
Public Const RiskIncrementPercent As Double = 0.5
Public Const RiskUnitPercent As Double = 1#

Public Const ScaleThresholdFactor As Double = 0.5

' number of ticks below/above a low/high to set a stop
Public Const StopBreakoutThresholdTicks = 1

Public Const StopIncrementFactor As Double = 0.5
Public Const SwingToBreakevenTicks As Long = 0
Public Const SwingToMoveStopTicks As Long = 0

Public Const TicksSwingToTrail As Long = 0

Public Const UseIntermediateStops As Boolean = False

Public Const ParamATRPeriods                       As String = "ATR Periods"

Public Const ParamBackstopBarLengthFactor           As String = "Backstop Bar Length Factor"
Public Const ParamBackstopMAPeriods                 As String = "Backstop MA Periods"
Public Const ParamBarLength                        As String = "Bar Length"
Public Const ParamBarUnit                           As String = "Bar Unit"
Public Const ParamBollingerCentreBandWidthTicks    As String = "Bollinger Centre Band Width Ticks"
Public Const ParamBollingerEdgeBandWidthTicks      As String = "Bollinger Edge Band Width Ticks"
Public Const ParamBollingerPeriods                 As String = "Bollinger Periods"
Public Const ParamBollingerMovingAverageType       As String = "Bollinger Moving Avg Type"
Public Const ParamBreakEvenBarLength               As String = "BreakEven Bar Length"
Public Const ParamBreakEvenBarThresholdTicks       As String = "BreakEven Bar Threshold Ticks"
Public Const ParamBreakEvenThresholdTicks          As String = "BreakEven Threshold Ticks"
Public Const ParamBreakoutThresholdTicks            As String = "Breakout Threshold Ticks"

Public Const ParamConfirmedBarsForBreakEven        As String = "Confirmed Bars For BreakEven"
Public Const ParamConfirmedBarsForExit             As String = "Confirmed Bars For Exit"
Public Const ParamConfirmedBarsForSwing            As String = "Confirmed Bars For Swing"
Public Const ParamConfirmedBarsForTrail            As String = "Confirmed Bars For Trail"

Public Const ParamEntryBreakoutThresholdTicks      As String = "Entry Breakout Threshold Ticks"
Public Const ParamEntryLimitOffsetTicks            As String = "Entry Limit Offset Ticks"

Public Const ParamIncludeBarsOutsideSession        As String = "Include Bars Outside Session"
Public Const ParamInitialStopFactor                 As String = "Initial Stop Factor"

Public Const ParamLongBarLengthFactor               As String = "Long Bar Length Factor"
Public Const ParamLongBollPeriods                   As String = "Long Bollinger Periods"
Public Const ParamLongMAPeriods                     As String = "Long MA Periods"

Public Const ParamMaxContraSwingPercent             As String = "Max Contra Swing Percent"
Public Const ParamMaxIncrements                     As String = "Max Increments"
Public Const ParamMaxInitialStopTicks              As String = "Max Initial Stop Ticks"
Public Const ParamMaxTradeSize                     As String = "Max Trade Size"
Public Const ParamMinimumSwingTicks                As String = "Minimum Swing Ticks"

Public Const ParamRetraceFromExtremes              As String = "Retrace From Extremes"
Public Const ParamRetracementStopPercent           As String = "Retracement Stop Percent"
Public Const ParamRetracementStopThresholdTicks    As String = "Retracement Stop Threshold Ticks"
Public Const ParamRewardToRiskRatio                 As String = "Reward To Risk Ratio"
Public Const ParamRiskIncrementPercent             As String = "Risk Increment Percent"
Public Const ParamRiskUnitPercent                  As String = "Risk Unit Percent"

Public Const ParamScaleThresholdFactor             As String = "Scale Threshold Factor"
Public Const ParamStopIncrementFactor               As String = "Stop Increment Factor"
Public Const ParamSwingToBreakevenTicks            As String = "Swing To Breakeven Ticks"
Public Const ParamSwingToMoveStopTicks              As String = "Swing To Move Stop Ticks"

Public Const ParamTicksSwingToTrail                 As String = "Ticks Swing To Trail"

Public Const ParamUseIntermediateStops              As String = "Use Intermediate Stops "


'@================================================================================
' Member variables
'@================================================================================

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

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




