VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Implemented Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

Event CancelMarketData( _
                    ByVal tickerid As Long, _
                    ByRef errcode As Long, _
                    ByRef errdesc As String)

Event CancelMarketDepth( _
                    ByVal tickerid As Long, _
                    ByVal forceCancel As Boolean, _
                    ByRef errcode As Long, _
                    ByRef errdesc As String)

'Event InputTickfileManager( _
'                    ByVal tickfile As TickFileManager, _
'                    ByVal pTicker As ticker, _
'                    ByRef errcode As Long, _
'                    ByRef errdesc As String)
Event NewTicker( _
                    ByRef pTicker As ticker, _
                    ByRef errcode As Long, _
                    ByRef errdesc As String)

Event RequestContract( _
                    ByVal contractSpec As contractSpecifier, _
                    ByVal pContracts As Contracts, _
                    ByRef errcode As Long, _
                    ByRef errdesc As String)

Event RequestMarketDataEx( _
                    ByVal pContractSpecifier As contractSpecifier, _
                    ByVal pTicker As ticker, _
                    ByVal DOMEventsRequired As DOMEvents, _
                    ByVal writeToTickFile As Boolean, _
                    ByVal includeMarketDepthInTickfile As Boolean, _
                    ByVal noVolumeAdjustments As Boolean, _
                    ByVal noImpliedTrades As Boolean, _
                    ByRef errcode As Long, _
                    ByRef errdesc As String)

Event RequestMarketDepth( _
                    ByVal pTicker As ticker, _
                    ByVal DOMEventsRequired As DOMEvents, _
                    ByVal writeToTickFile As Boolean, _
                    ByRef errcode As Long, _
                    ByRef errdesc As String)

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Sub CancelMarketData(ByVal tickerid As Long)
Dim errcode As Long
Dim errdesc As String

RaiseEvent CancelMarketData(tickerid, errcode, errdesc)

If errcode <> 0 Then
    err.Raise errcode, "TradeBuild.APIProxy::RequestContract", errdesc
End If
End Sub

Public Sub CancelMarketDepth(ByVal tickerid As Long, _
                            Optional ByVal forceCancel As Boolean = False)
Dim errcode As Long
Dim errdesc As String

RaiseEvent CancelMarketDepth(tickerid, forceCancel, errcode, errdesc)

If errcode <> 0 Then
    err.Raise errcode, "TradeBuild.APIProxy::RequestContract", errdesc
End If
End Sub

'Public Sub InputTickfileManager(ByVal tickfile As TickFileManager, _
'                                ByVal pTicker As ticker)
'Dim errcode As Long
'Dim errdesc As String
'
'RaiseEvent InputTickfileManager(tickfile, pTicker, errcode, errdesc)
'
'If errcode <> 0 Then
'    err.Raise errcode, "TradeBuild.APIProxy::RequestContract", errdesc
'End If
'End Sub

Public Function NewTicker() As ticker
Dim errcode As Long
Dim errdesc As String

RaiseEvent NewTicker(NewTicker, errcode, errdesc)

If errcode <> 0 Then
    err.Raise errcode, "TradeBuild.APIProxy::RequestContract", errdesc
End If
End Function

Public Sub RequestContract(ByVal contractSpec As contractSpecifier, _
                        ByVal pContracts As Contracts)
Dim errcode As Long
Dim errdesc As String

RaiseEvent RequestContract(contractSpec, pContracts, errcode, errdesc)

If errcode <> 0 Then
    err.Raise errcode, "TradeBuild.APIProxy::RequestContract", errdesc
End If
End Sub

Public Sub RequestMarketDataEx(ByVal pContractSpecifier As contractSpecifier, _
                            ByVal pTicker As ticker, _
                            Optional ByVal DOMEventsRequired As DOMEvents = DOMNoEvents, _
                            Optional ByVal writeToTickFile As Boolean = False, _
                            Optional ByVal includeMarketDepthInTickfile As Boolean = False, _
                            Optional ByVal noVolumeAdjustments As Boolean = False, _
                            Optional ByVal noImpliedTrades As Boolean = False)
Dim errcode As Long
Dim errdesc As String

RaiseEvent RequestMarketDataEx(pContractSpecifier, _
                            pTicker, _
                            DOMEventsRequired, _
                            writeToTickFile, _
                            includeMarketDepthInTickfile, _
                            noVolumeAdjustments, _
                            noImpliedTrades, _
                            errcode, _
                            errdesc)

If errcode <> 0 Then
    err.Raise errcode, "TradeBuild.APIProxy::RequestContract", errdesc
End If
End Sub

Public Sub RequestMarketDepth( _
                            ByVal pTicker As ticker, _
                            ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)
                            
Dim errcode As Long
Dim errdesc As String

RaiseEvent RequestMarketDepth(pTicker, _
                            DOMEventsRequired, _
                            writeToTickFile, _
                            errcode, _
                            errdesc)

If errcode <> 0 Then
    err.Raise errcode, "TradeBuild.APIProxy::RequestContract", errdesc
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================


