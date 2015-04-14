VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScaleStrategyFactory1"
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

Private Const ModuleName                            As String = "ScaleStrategyFactory1"

'@================================================================================
' Member variables
'@================================================================================

Private S                                           As New ScaleStrategy1Static

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
lParams.SetParameterValue ParamATRPeriods, ATRPeriods
lParams.SetParameterValue ParamIncludeBarsOutsideSession, IncludeBarsOutsideSession
lParams.SetParameterValue ParamInitialStopFactor, InitialStopFactor
lParams.SetParameterValue ParamMaxIncrements, MaxIncrements
lParams.SetParameterValue ParamRiskUnitPercent, RiskUnitPercent
lParams.SetParameterValue ParamRiskIncrementPercent, RiskIncrementPercent
lParams.SetParameterValue ParamScaleThresholdFactor, ScaleThresholdFactor
lParams.SetParameterValue ParamStopIncrementFactor, StopIncrementFactor

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
lBarLength = pParams.GetParameterValue(ParamBarLength)

Dim lBarUnit As TimePeriodUnits
lBarUnit = TimePeriodUnitsFromString(pParams.GetParameterValue(ParamBarUnit))

Dim lTimeframe As Timeframe
Set lTimeframe = pContext.AddTimeframe(GetTimePeriod(lBarLength, _
                                                lBarUnit), _
                                500, _
                                pParams.GetParameterValue(ParamIncludeBarsOutsideSession))

Dim params As Parameters

Set params = New Parameters
params.SetParameterValue "Periods", pParams.GetParameterValue(ParamATRPeriods)

ReDim linputvaluenames(0) As String
linputvaluenames(0) = "$default"
Set S.ATR = pContext.AddStudy("ATR", _
                        lTimeframe, _
                        linputvaluenames, _
                        pParams.GetParameterValue(ParamIncludeBarsOutsideSession), _
                        params, _
                        2)

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

S.InitialStopFactor = pParams.GetParameterValue(ParamInitialStopFactor)
S.MaxIncrements = pParams.GetParameterValue(ParamMaxIncrements)
S.RiskUnitPercent = pParams.GetParameterValue(ParamRiskUnitPercent)
S.RiskIncrementPercent = pParams.GetParameterValue(ParamRiskIncrementPercent)
S.ScaleThresholdFactor = pParams.GetParameterValue(ParamScaleThresholdFactor)
S.StopIncrementFactor = pParams.GetParameterValue(ParamStopIncrementFactor)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




