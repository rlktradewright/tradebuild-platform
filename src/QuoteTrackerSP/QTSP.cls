VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTSP"
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

'Public Enum TickTypes
'    None
'    bid
'    ask
'    last
'    volume
'    prevClose
'    high
'    low
'    openInterest
'End Enum

'@================================================================================
' Types
'@================================================================================

Public Type HistoricBar
    timestamp       As Date
    openValue       As Double
    highValue       As Double
    lowValue        As Double
    closeValue      As Double
    volume          As Long
End Type

Public Type QTHistoricalDataRequestQueueEntry
    symbol                  As String
    startDateTime           As Date
    endDateTime             As Date
    includeBarsOutsideSession   As Boolean
    barLength               As Long
    barUnits                As TimePeriodUnits
    rth                     As Boolean
End Type

'Public Type Tick
'    TickType    As TickTypes
'    timestamp   As Date
'    Price       As Double
'    size        As Long
'End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "QTSP26"
Private Const ModuleName                    As String = "QTSP"

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

'@================================================================================
' Helper Functions
'@================================================================================


