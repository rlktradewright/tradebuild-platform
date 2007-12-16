VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickUtils"
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

Public Enum TickTypes
    Bid
    Ask
    closePrice
    highPrice
    lowPrice
    marketDepth
    MarketDepthReset
    Trade
    volume
    openInterest
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type TickEvent
    source              As Object
    timestamp           As Date
    tickType            As TickTypes
    price               As Double
    size                As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TickUtils26"
Private Const ModuleName                    As String = "TickUtils"

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

Public Function CreateTickEncoder( _
                ByVal periodStartTime As Date, _
                ByVal tickSize As Double, _
                ByVal encodingFormat As String) As TickDataEncoder
If encodingFormat = TickfileFormatTradeBuildSQL Then
    Dim enc1 As TickDataEncoderv1
    Set enc1 = New TickDataEncoderv1
    Set CreateTickEncoder = enc1
    enc1.initialise periodStartTime, tickSize
End Function

'@================================================================================
' Helper Functions
'@================================================================================


