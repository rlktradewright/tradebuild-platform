VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScaleStrategy1Static"
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

Private Const ModuleName                            As String = "ScaleStrategy1Static"

'@================================================================================
' Member variables
'@================================================================================

Public ATR                                          As ATR
Public MaxIncrements                                As Long

' the percentage of account balance when 1 unit of contracts
' changes price by 1 ATR
Public RiskUnitPercent                              As Double
                                            
' the amount by which risk is to be increased at each scaling point
Public RiskIncrementPercent                         As Double

' the number of ATRs to the next scaling point
Public ScaleThresholdFactor                         As Double

' the number of ATRs to raise the stop losses at each scaling point
Public StopIncrementFactor                          As Double

' the number of ATR to set the initial stop loss for each order
Public InitialStopFactor                            As Double

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




