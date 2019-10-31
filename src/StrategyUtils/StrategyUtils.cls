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
    EntryOrderTypeNone = OrderTypes.OrderTypeNone
    EntryOrderTypeMarket = OrderTypes.OrderTypeMarket
    EntryOrderTypeMarketOnOpen = OrderTypes.OrderTypeMarketOnOpen
    EntryOrderTypeMarketOnClose = OrderTypes.OrderTypeMarketOnClose
    EntryOrderTypeMarketIfTouched = OrderTypes.OrderTypeMarketIfTouched
    EntryOrderTypeMarketToLimit = OrderTypes.OrderTypeMarketToLimit
    EntryOrderTypeLimit = OrderTypes.OrderTypeLimit
    EntryOrderTypeLimitOnOpen = OrderTypes.OrderTypeLimitOnOpen
    EntryOrderTypeLimitOnClose = OrderTypes.OrderTypeLimitOnClose
    EntryOrderTypeLimitIfTouched = OrderTypes.OrderTypeLimitIfTouched
    EntryOrderTypeStop = OrderTypes.OrderTypeStop
    EntryOrderTypeStopLimit = OrderTypes.OrderTypeStopLimit
    EntryOrderTypeTrailStop = OrderTypes.OrderTypeTrail
    EntryOrderTypeTrailStopLimit = OrderTypes.OrderTypeTrailLimit
End Enum

Public Enum OrderTIFs
    OrderTIFNone = 0
    OrderTIFDay = 1
    OrderTIFGoodTillCancelled = 2
    OrderTIFImmediateOrCancel = 4
End Enum

Public Enum StopLossOrderTypes
    StopLossOrderTypeNone = OrderTypes.OrderTypeNone
    StopLossOrderTypeStop = OrderTypes.OrderTypeStop
    StopLossOrderTypeStopLimit = OrderTypes.OrderTypeStopLimit
    StopLossOrderTypeTrailStop = OrderTypes.OrderTypeTrail
    StopLossOrderTypeTrailStopLimit = OrderTypes.OrderTypeTrailLimit
End Enum

Public Enum TargetOrderTypes
    TargetOrderTypeNone = OrderTypes.OrderTypeNone
    TargetOrderTypeLimit = OrderTypes.OrderTypeLimit
    TargetOrderTypeLimitIfTouched = OrderTypes.OrderTypeLimitIfTouched
    TargetOrderTypeMarketIfTouched = OrderTypes.OrderTypeMarketIfTouched
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

Set CreateStrategyRunner = gCreateStrategyRunner(pStrategyHost)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


