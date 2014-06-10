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

Implements IStrategyFactory

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

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStrategyFactory Interface Members
'@================================================================================

Private Function IStrategyFactory_CreateStrategy( _
                ByVal pTradingContext As TradingContext) As IPositionMgmtStrategy
Const ProcName As String = "IStrategyFactory_CreateStrategy"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IStrategyFactory_Initialise( _
                ByVal pContext As InitialisationContext)
Const ProcName As String = "IStrategyFactory_Initialise"
On Error GoTo Err

Dim barsID As String
Dim studyID As String
Dim params As cParameters

Dim lTFD As TimeframeDesignator
Set lTFD = pContext.AddTimeframe(GetTimePeriod(mTimePeriod, mTimePeriodUnit), barsID)

Set params = New Parameters
params.setParam "Periods", mATRPeriods

ReDim lInputValueNames(0) As String
lInputValueNames(0) = "$default"
Set mATRStudy = pContext.AddStudy("AverageTrueRange", _
                        pBaseStudy, _
                        lInputValueNames, _
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




