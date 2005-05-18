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

Event AddListener(ByVal newListener As IListener, _
                ByRef valueType As Long, _
                ByRef listenerKey As String)

Event Application( _
                ByRef data As Object)

Event CancelMarketDepth(ByVal forceCancel As Boolean)

Event GetAskPrice(ByRef price As Double)

Event GetAskSize(ByRef Size As Long)

Event GetBidPrice(ByRef price As Double)

Event GetBidSize(ByRef Size As Long)

Event GetChange(ByRef price As Double)

Event GetChangePercent(ByRef value As Double)

Event GetClosePrice(ByRef price As Double)

Event GetContract(ByRef contract As contract)

Event GetDescription(ByRef value As String)

Event GetHighPrice(ByRef price As Double)

Event GetId(ByRef ID As Long)

Event GetKey(ByRef key As String)

Event GetLowPrice(ByRef price As Double)

Event GetPosition(ByRef position As Long)

Event GetTickfileSpecifiers(ByRef TickfileSpecifiers() As TradeBuild.TickfileSpecifier)

Event GetState(ByRef State As TickerStateCodes)

Event GetTradePrice(ByRef price As Double)

Event GetTradeSize(ByRef Size As Long)

Event GetVolume(ByRef Volume As Long)

Event PauseTicker()

Event RemoveListener(ByVal key As String)

Event RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)

Event selectTickFiles()

Event SetTickfileSpecifiers(ByRef value() As TradeBuild.TickfileSpecifier)

Event SkipTicker()

Event StartTicker(ByVal pContractSpecifier As ContractSpecifier, _
                        ByVal DOMEventsRequired As DOMEvents, _
                        ByVal writeToTickFile As Boolean, _
                        ByVal includeMarketDepthInTickfile As Boolean, _
                        ByVal noVolumeAdjustments As Boolean, _
                        ByVal noImpliedTrades As Boolean)

Event StopTicker()

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

Friend Property Get ClosePrice() As Double
RaiseEvent GetClosePrice(ClosePrice)
End Property

Friend Property Get contract() As contract
RaiseEvent GetContract(contract)
End Property

Friend Property Get Description() As String
RaiseEvent GetDescription(Description)
End Property

Friend Property Get HighPrice() As Double
RaiseEvent GetHighPrice(HighPrice)
End Property

Friend Property Get ID() As Long
RaiseEvent GetId(ID)
End Property

Friend Property Get key() As String
RaiseEvent GetKey(key)
End Property

Friend Property Get LowPrice() As Double
RaiseEvent GetLowPrice(LowPrice)
End Property

Friend Property Get position() As Long
RaiseEvent GetPosition(position)
End Property

Friend Property Get State() As TickerStateCodes
RaiseEvent GetState(State)
End Property

Friend Property Let TickfileSpecifiers(value() As TradeBuild.TickfileSpecifier)
RaiseEvent SetTickfileSpecifiers(value)
End Property

Friend Property Get TickfileSpecifiers() As TradeBuild.TickfileSpecifier()
RaiseEvent GetTickfileSpecifiers(TickfileSpecifiers)
End Property

Friend Property Get TradePrice() As Double
RaiseEvent GetTradePrice(TradePrice)
End Property

Friend Property Get TradeSize() As Long
RaiseEvent GetTradeSize(TradeSize)
End Property

Friend Property Get Volume() As Long
RaiseEvent GetVolume(Volume)
End Property

'================================================================================
' Methods
'================================================================================

Friend Function AddListener(ByVal newListener As IListener, _
                        ByVal valueType As Long) As String
RaiseEvent AddListener(newListener, valueType, AddListener)
End Function

Friend Sub FireApplicationEvent(ByRef data As Object)
RaiseEvent Application(data)
End Sub

Friend Sub CancelMarketDepth(Optional ByVal forceCancel As Boolean = False)
RaiseEvent CancelMarketDepth(forceCancel)
End Sub

Friend Sub PauseTicker()
RaiseEvent PauseTicker
End Sub

Friend Sub RemoveListener(ByVal key As String)
RaiseEvent RemoveListener(key)
End Sub

Friend Sub selectTickFiles()
RaiseEvent selectTickFiles
End Sub

Friend Sub SkipTicker()
RaiseEvent SkipTicker
End Sub

Friend Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)
RaiseEvent RequestMarketDepth(DOMEventsRequired, _
                        writeToTickFile)
End Sub

Friend Sub StartTicker(ByVal pContractSpecifier As ContractSpecifier, _
                        Optional ByVal DOMEventsRequired As DOMEvents = DOMNoEvents, _
                        Optional ByVal writeToTickFile As Boolean = False, _
                        Optional ByVal includeMarketDepthInTickfile As Boolean = False, _
                        Optional ByVal noVolumeAdjustments As Boolean = False, _
                        Optional ByVal noImpliedTrades As Boolean = False)
RaiseEvent StartTicker(pContractSpecifier, _
                        DOMEventsRequired, _
                        writeToTickFile, _
                        includeMarketDepthInTickfile, _
                        noVolumeAdjustments, _
                        noImpliedTrades)
End Sub

Friend Sub StopTicker()
RaiseEvent StopTicker
End Sub

'================================================================================
' Helper Functions
'================================================================================




