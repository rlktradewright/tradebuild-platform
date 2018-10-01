VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderSpecification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "BracketOrderSpecification"

'@================================================================================
' Member variables
'@================================================================================

Public Action                                       As OrderActions
Public Quantity                                     As Long
Public EntryType                                    As BracketEntryTypes
Public EntryPrice                                   As Double
Public EntryOffset                                  As Long
Public EntryTriggerPrice                            As Double
Public StopLossType                                 As BracketStopLossTypes
Public StopLossPrice                                As Double
Public StopLossOffset                               As Long
Public StopLossTriggerPrice                         As Double
Public TargetType                                   As BracketTargetTypes
Public TargetPrice                                  As Double
Public TargetOffset                                 As Long
Public TargetTriggerPrice                           As Double
Public EntryTIF                                     As OrderTIFs
Public StopTIF                                      As OrderTIFs
Public TargetTIF                                    As OrderTIFs
Public CancelPrice                                  As Double
Public CancelAfter                                  As Long

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




