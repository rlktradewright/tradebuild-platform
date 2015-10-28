VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopStrategy5Static"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "StopStrategy5Static"

'@================================================================================
' Member variables
'@================================================================================

Public LongTimeframe As ResourceIdentifier

Public StopTimeframe As ResourceIdentifier
Public StopMA As ResourceIdentifier
Public Boll As ResourceIdentifier

Public LongBoll As ResourceIdentifier

Public BreakevenTimeframe As ResourceIdentifier

Public LongMA As ResourceIdentifier

Public BreakEvenThresholdTicks As Long
Public BreakevenBarThresholdTicks As Integer

Public BreakoutThresholdTicks As Integer

Public ConfirmedBarsForBreakEven As Long
Public ConfirmedBarsForSwing As Long
Public ConfirmedBarsForTrailing As Long
Public ConfirmedBarsForExit As Long

Public MaxContraSwingFactor As Double

Public RetraceFromExtremes As Boolean
Public RetracementStopFactor As Double
Public RetracementStopThresholdTicks As Long
Public RewardToRiskRatio As Single

Public SwingToMoveStopTicks As Long

Public TicksSwingToTrail As Long

Public UseIntermediateStops As Boolean

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

'@================================================================================
' Helper Functions
'@================================================================================




