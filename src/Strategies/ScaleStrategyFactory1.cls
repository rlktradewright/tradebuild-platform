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

Private Sub IPosnMgmtStrtgyFactory_DefineDefaultParameters()
Const ProcName As String = "IPosnMgmtStrtgyFactory_DefineDefaultParameters"
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

Dim lTimeframe As ResourceIdentifier
Set lTimeframe = DeclareTimeframe(lBarLength, _
                                lBarUnit, _
                                500, _
                                GetParameterValue(ParamIncludeBarsOutsideSession))

Set S.ATR = DeclareStudy( _
                        "ATR", _
                        lTimeframe, _
                        "Periods=" & GetParameterValue(ParamATRPeriods), _
                        GetParameterValue(ParamIncludeBarsOutsideSession), _
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

Private Sub processParameters()
Const ProcName As String = "processParameters"
On Error GoTo Err

S.InitialStopFactor = GetParameterValue(ParamInitialStopFactor)
S.MaxIncrements = GetParameterValue(ParamMaxIncrements)
S.RiskUnitPercent = GetParameterValue(ParamRiskUnitPercent)
S.RiskIncrementPercent = GetParameterValue(ParamRiskIncrementPercent)
S.ScaleThresholdFactor = GetParameterValue(ParamScaleThresholdFactor)
S.StopIncrementFactor = GetParameterValue(ParamStopIncrementFactor)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




