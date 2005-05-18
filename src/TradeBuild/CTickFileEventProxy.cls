VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickFileEventProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Ask( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event Bid( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event High( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event Low( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event PreviousClose( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double)

Event ReplayCompleted( _
                ByVal tickerID As Long)

Event ResetMarketDepth( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date)

Event StartingTickfile( _
                ByVal tickerID As Long, _
                ByVal pContract As contract, _
                ByVal replaySpeed As Long, _
                ByVal rewrite As Boolean, _
                ByVal noVolumeAdjustments As Boolean, _
                ByVal noImpliedTrades As Boolean, _
                ByVal DOMEventsRequired As DOMEvents)
                
Event Trade( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)

Event UpdateMktDepth( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
                
Event UpdateMktDepthL2( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)

Event Volume( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal Size As Long)

Friend Sub Ask( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)
RaiseEvent Ask( _
                tickerID, _
                timestamp, _
                price, _
                Size)
End Sub

Friend Sub Bid( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)
RaiseEvent Bid( _
                tickerID, _
                timestamp, _
                price, _
                Size)
End Sub

Friend Sub High( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double)
RaiseEvent High( _
                tickerID, _
                timestamp, _
                price)
End Sub

Friend Sub Low( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double)
RaiseEvent Low( _
                tickerID, _
                timestamp, _
                price)
End Sub

Friend Sub nextTickfile( _
                ByVal tickerID As Long, _
                ByVal pContract As contract, _
                ByVal replaySpeed As Long, _
                ByVal rewrite As Boolean, _
                ByVal noVolumeAdjustments As Boolean, _
                ByVal noImpliedTrades As Boolean, _
                ByVal DOMEventsRequired As DOMEvents)
RaiseEvent nextTickfile( _
                tickerID, _
                pContract, _
                replaySpeed, _
                rewrite, _
                noVolumeAdjustments, _
                noImpliedTrades, _
                DOMEventsRequired)
End Sub

Friend Sub PreviousClose( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double)
RaiseEvent PreviousClose( _
                tickerID, _
                timestamp, _
                price)
End Sub

Friend Sub ReplayCompleted(ByVal tickerID As Long)
RaiseEvent ReplayCompleted(tickerID)
End Sub

Friend Sub ResetMarketDepth( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date)
RaiseEvent ResetMarketDepth( _
                tickerID, _
                timestamp)
End Sub

Friend Sub Trade( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal Size As Long)
RaiseEvent Trade( _
                tickerID, _
                timestamp, _
                price, _
                Size)
End Sub

Friend Sub UpdateMktDepth( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
RaiseEvent UpdateMktDepth( _
                tickerID, _
                timestamp, _
                position, _
                operation, _
                side, _
                price, _
                Size)

End Sub
                
Friend Sub UpdateMktDepthL2( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
RaiseEvent UpdateMktDepthL2( _
                tickerID, _
                timestamp, _
                position, _
                marketMaker, _
                operation, _
                side, _
                price, _
                Size)

End Sub

Friend Sub Volume( _
                ByVal tickerID As Long, _
                ByVal timestamp As Date, _
                ByVal Size As Long)
RaiseEvent Volume( _
                tickerID, _
                timestamp, _
                Size)
End Sub



