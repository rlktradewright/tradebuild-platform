VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickerProxyEventProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Event Application( _
                ByVal data As Variant)

Event CancelMarketDepth(ByVal forceCancel As Boolean)

Event GetAskPrice(ByRef price As Double)

Event GetAskSize(ByRef Size As Long)

Event GetBidPrice(ByRef price As Double)

Event GetBidSize(ByRef Size As Long)

Event GetChange(ByRef price As Double)

Event GetChangePercent(ByRef value As Double)

Event GetClosePrice(ByRef price As Double)

Event GetContract(ByRef Contract As Contract)

Event GetDescription(ByRef value As String)

Event GetDOMEventsRequired(ByRef eventsRequired As DOMEvents)

Event GetHighPrice(ByRef price As Double)

Event GetId(ByRef id As Long)

Event GetIncludeMarketDepthInTickfile(ByRef value As Boolean)

Event GetKey(ByRef key As String)

Event GetLowPrice(ByRef price As Double)

Event GetNoImpliedTrades(ByRef value As Boolean)

Event GetNoVolumeAdjustments(ByRef value As Boolean)

Event GetOutputTickfileFormat(ByRef outputFormat As String)

Event GetOutputTickFilePath(ByRef path As String)

Event GetPosition(ByRef position As Long)

Event GetPriceFormatString(ByRef formatString As String)

Event GetReplayingTickfile(ByRef value As Boolean)

Event GetState(ByRef State As TickerStateCodes)

Event GetTimeframes(ByRef pTimeframes As TradeBuild.timeframes)

Event GetTimestamp(ByRef timestamp As Date)

Event GetTradePrice(ByRef price As Double)

Event GetTradeSize(ByRef Size As Long)

Event GetVolume(ByRef volume As Long)

Event RemoveListener(ByVal key As String)

Event RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickfile As Boolean)

Event StartTicker(ByVal pContractSpecifier As contractSpecifier)

Event StopTicker()

Event GetWriteToTickFile(ByRef value As Boolean)

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

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get AskPrice() As Double
RaiseEvent GetAskPrice(AskPrice)
End Property

Friend Property Get AskSize() As Long
RaiseEvent GetAskSize(AskSize)
End Property

Friend Property Get BidPrice() As Double
RaiseEvent GetBidPrice(BidPrice)
End Property

Friend Property Get bidSize() As Long
RaiseEvent GetBidSize(bidSize)
End Property

Friend Property Get Change() As Double
RaiseEvent GetChange(Change)
End Property

Friend Property Get ChangePercent() As Double
RaiseEvent GetChangePercent(ChangePercent)
End Property

Friend Property Get closePrice() As Double
RaiseEvent GetClosePrice(closePrice)
End Property

Friend Property Get Contract() As Contract
RaiseEvent GetContract(Contract)
End Property

Friend Property Get Description() As String
RaiseEvent GetDescription(Description)
End Property

Friend Property Get DOMEventsRequired() As DOMEvents
RaiseEvent GetDOMEventsRequired(DOMEventsRequired)
End Property

Friend Property Get highPrice() As Double
RaiseEvent GetHighPrice(highPrice)
End Property

Friend Property Get id() As Long
RaiseEvent GetId(id)
End Property

Friend Property Get includeMarketDepthInTickfile() As Boolean
RaiseEvent GetIncludeMarketDepthInTickfile(includeMarketDepthInTickfile)
End Property

Friend Property Get key() As String
RaiseEvent GetKey(key)
End Property

Friend Property Get lowPrice() As Double
RaiseEvent GetLowPrice(lowPrice)
End Property

Friend Property Get noImpliedTrades() As Boolean
RaiseEvent GetNoImpliedTrades(noImpliedTrades)
End Property

Friend Property Get noVolumeAdjustments() As Boolean
RaiseEvent GetNoVolumeAdjustments(noVolumeAdjustments)
End Property

Friend Property Get outputTickfileFormat() As String
RaiseEvent GetOutputTickfileFormat(outputTickfileFormat)
End Property

Friend Property Get outputTickfilePath() As String
RaiseEvent GetOutputTickFilePath(outputTickfilePath)
End Property

Friend Property Get position() As Long
RaiseEvent GetPosition(position)
End Property

Friend Property Get priceFormatString() As String
RaiseEvent GetPriceFormatString(priceFormatString)
End Property

Friend Property Get replayingTickfile() As Boolean
RaiseEvent GetReplayingTickfile(replayingTickfile)
End Property

Friend Property Get State() As TickerStateCodes
RaiseEvent GetState(State)
End Property

Friend Property Get timeframes() As timeframes
RaiseEvent GetTimeframes(timeframes)
End Property

Friend Property Get timestamp() As Date
RaiseEvent GetTimestamp(timestamp)
End Property

Friend Property Get TradePrice() As Double
RaiseEvent GetTradePrice(TradePrice)
End Property

Friend Property Get TradeSize() As Long
RaiseEvent GetTradeSize(TradeSize)
End Property

Friend Property Get volume() As Long
RaiseEvent GetVolume(volume)
End Property

Friend Property Get writeToTickfile() As Boolean
RaiseEvent GetWriteToTickFile(writeToTickfile)
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub FireApplicationEvent(ByRef data As Variant)
RaiseEvent Application(data)
End Sub

Friend Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
RaiseEvent CancelMarketDepth(forceCancel)
End Sub

Friend Sub RemoveListener(ByVal key As String)
RaiseEvent RemoveListener(key)
End Sub

Friend Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickfile As Boolean)
RaiseEvent RequestMarketDepth(DOMEventsRequired, _
                        writeToTickfile)
End Sub

Friend Sub StartTicker(ByVal pContractSpecifier As contractSpecifier)
RaiseEvent StartTicker(pContractSpecifier)
End Sub

Friend Sub StopTicker()
RaiseEvent StopTicker
End Sub

'================================================================================
' Helper Functions
'================================================================================




