VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBTWSSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'Public Type Execution
'    ExecId         As String
'    time           As Date
'    AccountIdentifier As String
'    ContractSpecifier As ContractSpecifier
'    FillingExchange As String
'    action         As OrderActions
'    Quantity       As Long
'    price          As Double
'    OrderBrokerId  As String
'    OrderTradeBuildId  As String
'    IsLiquidation  As Boolean
'End Type

Public Type HistoricBar
    timestamp       As Date
    openValue       As Double
    highValue       As Double
    lowValue        As Double
    closeValue      As Double
    volume          As Long
    WAP             As Double
    hasGaps         As Boolean
    tickVolume      As Long
End Type

Public Type OrderIdMapEntry
    order                   As IOrder
    OrderSubmitter          As OrderSubmitter
End Type

Public Type TWSContractDetailsRequestQueueEntry
    contractSpec                As ContractSpecifier
    contractWrappers            As Collection
    requester                   As ContractDetailsRequester
End Type

Public Type TWSHistoricalDataRequestQueueEntry
    id                      As Long
    Contract                As Contract
    endDateTime             As String
    barSizeSetting          As String
    duration                As String
    whatToShow              As String
End Type


