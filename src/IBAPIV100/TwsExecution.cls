VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsExecution"
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

Private Const ModuleName                            As String = "TwsExecution"

'@================================================================================
' Member variables
'@================================================================================

Public OrderId                                    As Long
Public ClientID                                   As Long
Public ExecId                                     As String
Public Time                                       As String
Public AcctNumber                                 As String
Public Exchange                                   As String
Public Side                                       As String
Public Shares                                     As Long
Public Price                                      As Double
Public PermId                                     As Long
Public Liquidation                                As Boolean
Public CumQty                                     As Long
Public AvgPrice                                   As Double
Public OrderRef                                   As String
Public EvRule                                     As String
Public EvMultiplier                               As Double
Public ModelCode                                    As String
Public LastLiquidity                                As TwsLiquidityTypes

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




