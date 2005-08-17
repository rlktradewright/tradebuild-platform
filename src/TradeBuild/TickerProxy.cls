VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickerProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'   25/01/05    First version created.
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mEventProxy As TickerProxyEventProxy

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mEventProxy = New TickerProxyEventProxy
End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get AskPrice() As Double
AskPrice = mEventProxy.AskPrice
End Property

Public Property Get AskSize() As Long
AskSize = mEventProxy.AskSize
End Property

Public Property Get BidPrice() As Double
BidPrice = mEventProxy.BidPrice
End Property

Public Property Get bidSize() As Long
bidSize = mEventProxy.bidSize
End Property

Public Property Get Change() As Double
Change = mEventProxy.Change
End Property

Public Property Get ChangePercent() As Double
ChangePercent = mEventProxy.ChangePercent
End Property

Public Property Get closePrice() As Double
closePrice = mEventProxy.closePrice
End Property

Public Property Get Contract() As Contract
Contract = mEventProxy.Contract
End Property

Public Property Get Description() As String
Description = mEventProxy.Description
End Property

Public Property Get DOMEventsRequired() As DOMEvents
DOMEventsRequired = mEventProxy.DOMEventsRequired
End Property

Friend Property Get EventProxy() As TickerProxyEventProxy
Set EventProxy = mEventProxy
End Property

Public Property Get highPrice() As Double
highPrice = mEventProxy.highPrice
End Property

Public Property Get id() As Long
id = mEventProxy.id
End Property

Public Property Get includeMarketDepthInTickfile() As Boolean
includeMarketDepthInTickfile = mEventProxy.includeMarketDepthInTickfile
End Property

Public Property Get key() As String
key = mEventProxy.key
End Property

Public Property Get lowPrice() As Double
lowPrice = mEventProxy.lowPrice
End Property

Public Property Get noImpliedTrades() As Boolean
noImpliedTrades = mEventProxy.noImpliedTrades
End Property

Public Property Get noVolumeAdjustments() As Boolean
noVolumeAdjustments = mEventProxy.noVolumeAdjustments
End Property

Public Property Get outputTickfileFormat() As String
outputTickfileFormat = mEventProxy.outputTickfileFormat
End Property

Public Property Get outputTickfilePath() As String
outputTickfilePath = mEventProxy.outputTickfilePath
End Property

Public Property Get position() As Long
position = mEventProxy.position
End Property

Public Property Get priceFormatString() As String
priceFormatString = mEventProxy.priceFormatString
End Property

Public Property Get replayingTickfile() As Boolean
replayingTickfile = mEventProxy.replayingTickfile
End Property

Public Property Get State() As TickerStateCodes
State = mEventProxy.State
End Property

Public Property Get timeframes() As TradeBuild.timeframes
timeframes = mEventProxy.timeframes
End Property

Public Property Get timestamp() As Date
timestamp = mEventProxy.timestamp
End Property

Public Property Get TradePrice() As Double
TradePrice = mEventProxy.TradePrice
End Property

Public Property Get TradeSize() As Long
TradeSize = mEventProxy.TradeSize
End Property

Public Property Get volume() As Long
volume = mEventProxy.volume
End Property

Public Property Let writeToTickfile(ByVal value As Boolean)
writeToTickfile = mEventProxy.writeToTickfile
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddListener(ByVal newListener As IListener, _
                        ByVal valueType As Long) As String
AddListener = mEventProxy.AddListener(newListener, valueType)
End Function

Public Sub FireApplicationEvent(ByRef data As Variant)
mEventProxy.FireApplicationEvent data
End Sub

Public Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
mEventProxy.CancelMarketDepth forceCancel
End Sub

Public Sub RemoveListener(ByVal key As String)
mEventProxy.RemoveListener key
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickfile As Boolean)
mEventProxy.RequestMarketDepth DOMEventsRequired, _
                        writeToTickfile
End Sub

Public Sub StartTicker(ByVal pContractSpecifier As contractSpecifier)
mEventProxy.StartTicker pContractSpecifier
End Sub

Public Sub StopTicker()
mEventProxy.StopTicker
End Sub

'================================================================================
' Helper Functions
'================================================================================




