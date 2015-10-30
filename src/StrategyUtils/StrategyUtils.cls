VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StrategyUtils"
Attribute VB_GlobalNameSpace = True
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

Public Enum EntryOrderTypes
    EntryOrderTypeNone = 0
    EntryOrderTypeMarket = 1
    EntryOrderTypeMarketOnOpen = &H100000
    EntryOrderTypeMarketOnClose = 2
    EntryOrderTypeMarketIfTouched = &H20000
    EntryOrderTypeMarketToLimit = &H400&
    EntryOrderTypeBid = &H1000000
    EntryOrderTypeAsk = &H2000000
    EntryOrderTypeLast = &H4000000
    EntryOrderTypeLimit = 4
    EntryOrderTypeLimitOnOpen = &H200000
    EntryOrderTypeLimitOnClose = 8
    EntryOrderTypeLimitIfTouched = &H10000
    EntryOrderTypeStop = &H20&
    EntryOrderTypeStopLimit = &H40&
End Enum

Public Enum OrderTIFs
    OrderTIFNone = 0
    OrderTIFDay = 1
    OrderTIFGoodTillCancelled = 2
    OrderTIFImmediateOrCancel = 4
End Enum

Public Enum StopLossOrderTypes
    StopLossOrderTypeNone = 0
    StopLossOrderTypeStop = &H20&
    StopLossOrderTypeStopLimit = &H40&
    StopLossOrderTypeBid = &H1000000
    StopLossOrderTypeAsk = &H2000000
    StopLossOrderTypeLast = &H4000000
    StopLossOrderTypeAuto = &H8000000
End Enum

Public Enum TargetOrderTypes
    TargetOrderTypeNone = 0
    TargetOrderTypeLimit = 4
    TargetOrderTypeLimitIfTouched = &H10000
    TargetOrderTypeMarketIfTouched = &H20000
    TargetOrderTypeBid = &H1000000
    TargetOrderTypeAsk = &H2000000
    TargetOrderTypeLast = &H4000000
    TargetOrderTypeAuto = &H8000000
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "StrategyUtils"

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

Public Function CreateStrategyRunner( _
                ByVal pStrategyHost As IStrategyHost) As StrategyRunner
Const ProcName As String = "CreateStrategyRunner"
On Error GoTo Err

AssertArgument Not pStrategyHost Is Nothing, "pStrategyHost is Nothing"

Set CreateStrategyRunner = New StrategyRunner
CreateStrategyRunner.Initialise pStrategyHost

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


