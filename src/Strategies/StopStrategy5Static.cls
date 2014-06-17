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

Public Timeframe As Timeframe
Public LongTimeframe As Timeframe
Public Swing As Swing
Public MACDStudy As MACD
Public ListenerKey As String

Public StopTimeframe As Timeframe
Public StopMA As EMA
Public Boll As BollingerBands

Public LongBoll As BollingerBands

Public BreakevenTimeframe As Timeframe

Public LongMA As SMA

Public BackstopMABarLength As Integer
Public BackstopMAOffset As Integer

Public BreakevenBarLength As Integer
Public BreakEvenThresholdTicks As Long
Public BreakevenBarThresholdTicks As Integer

Public BreakoutThreshold As Integer

Public ConfirmedCountForBreakEven As Long
Public ConfirmedCountForSwing As Long
Public ConfirmedCountForTrailing As Long
Public ConfirmedCountForExit As Long

Public RetraceFromExtremes As Boolean
Public RetracementStopFactor As Single
Public RetracementStopThresholdTicks As Long

Public TicksSwingToMoveStop As Long

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




