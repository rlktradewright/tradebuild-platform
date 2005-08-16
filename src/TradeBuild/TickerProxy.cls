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

Friend Property Get EventProxy() As TickerProxyEventProxy
Set EventProxy = mEventProxy
End Property

Public Property Get highPrice() As Double
highPrice = mEventProxy.highPrice
End Property

Public Property Get id() As Long
id = mEventProxy.id
End Property

Public Property Get key() As String
key = mEventProxy.key
End Property

Public Property Get lowPrice() As Double
lowPrice = mEventProxy.lowPrice
End Property

Public Property Get position() As Long
position = mEventProxy.position
End Property

Public Property Get State() As TickerStateCodes
State = mEventProxy.State
End Property

Public Property Let TickfileSpecifiers(value() As TradeBuild.TickfileSpecifier)
mEventProxy.TickfileSpecifiers = value
End Property

Public Property Get TickfileSpecifiers() As TradeBuild.TickfileSpecifier()
TickfileSpecifiers = mEventProxy.TickfileSpecifiers
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

Public Sub PauseTicker()
mEventProxy.PauseTicker
End Sub

Public Sub RemoveListener(ByVal key As String)
mEventProxy.RemoveListener key
End Sub

Public Sub selectTickFiles()
mEventProxy.selectTickFiles
End Sub

Public Sub SkipTicker()
mEventProxy.SkipTicker
End Sub

Public Sub RequestMarketDepth(ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)
mEventProxy.RequestMarketDepth DOMEventsRequired, _
                        writeToTickFile
End Sub

Public Sub StartTicker(ByVal pContractSpecifier As contractSpecifier, _
                        Optional ByVal DOMEventsRequired As DOMEvents = DOMNoEvents, _
                        Optional ByVal writeToTickFile As Boolean = False, _
                        Optional ByVal includeMarketDepthInTickfile As Boolean = False, _
                        Optional ByVal noVolumeAdjustments As Boolean = False, _
                        Optional ByVal noImpliedTrades As Boolean = False)
mEventProxy.StartTicker pContractSpecifier, _
                        DOMEventsRequired, _
                        writeToTickFile, _
                        includeMarketDepthInTickfile, _
                        noVolumeAdjustments, _
                        noImpliedTrades
End Sub

Public Sub StopTicker()
mEventProxy.StopTicker
End Sub

'================================================================================
' Helper Functions
'================================================================================




