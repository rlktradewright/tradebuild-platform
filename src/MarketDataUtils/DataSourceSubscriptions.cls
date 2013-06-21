VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSourceSubscriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DataSourceSubscriptions"

'@================================================================================
' Member variables
'@================================================================================

Private mDataStreamRef                              As WeakReference

Private mDataSources                                As New EnumerableCollection
Private mNumberOfMarketDataSubscriptions            As Long
Private mNumberOfMarketDepthSubscriptions           As Long

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

Friend Property Get NumberOfMarketDataSubscriptions() As Long
NumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions
End Property

Friend Property Get NumberOfMarketDepthSubscriptions() As Long
NumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions
End Property

Friend Property Get NumberOfSubscribers() As Long
NumberOfSubscribers = mDataSources.Count
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "AddDataSource"
On Error GoTo Err

Dim lSubscr As New DataSourceSubscription
lSubscr.DataSource = pDataSource
mDataSources.Add lSubscr, GetObjectKey(pDataSource)

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddMarketDataSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "AddMarketDataSubscription"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription

Set lSubscr = getSubscription(pDataSource)

If Not lSubscr.SubscribedToMarketData Then
    lSubscr.SubscribedToMarketData = True
    mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions + 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddMarketDepthSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "AddMarketDepthSubscription"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription

Set lSubscr = getSubscription(pDataSource)

If Not lSubscr.SubscribedToMarketDepth Then
    lSubscr.SubscribedToMarketDepth = True
    mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions + 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DOMReset()
Const ProcName As String = "DOMReset"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription

For Each lSubscr In mDataSources
    If lSubscr.SubscribedToMarketDepth Then fireResetMarketDepth lSubscr.DataSource
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DOMSetCell( _
                ByVal pSide As DOMSides, _
                ByVal pPrice As Double, _
                ByVal pSize As Long)
Const ProcName As String = "DOMSetCell"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription

For Each lSubscr In mDataSources
    If lSubscr.SubscribedToMarketDepth Then fireSetMarketDepthCell lSubscr.DataSource, pSide, pPrice, pSize
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Friend Sub Initialise(ByVal pDataStream As MarketDataStream)
Set mDataStreamRef = CreateWeakReference(pDataStream)
End Sub

Friend Sub NotifyContract(ByVal pContract As IContract)
Const ProcName As String = "NotifyContract"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription

For Each lSubscr In mDataSources
    lSubscr.DataSource.Contract = pContract
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

If mDataSources.Count = 0 Then Exit Sub

ReDim lSubscrs(mDataSources.Count - 1) As DataSourceSubscription
Dim lSubscr As DataSourceSubscription
Dim i As Long

For Each lSubscr In mDataSources
    Set lSubscrs(i) = lSubscr
Next

Dim lVar As Variant
For Each lVar In lSubscrs
    Set lSubscr = lVar
    lSubscr.DataSource.SetError pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDataError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDataError"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription

For Each lSubscr In mDataSources
    If lSubscr.SubscribedToMarketData Then fireError lSubscr.DataSource, pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDepthError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription

For Each lSubscr In mDataSources
    If lSubscr.SubscribedToMarketDepth Then fireError lSubscr.DataSource, pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim lGotQuote As Boolean
Dim lQuote As Quote
Dim lGotPriceChange As Boolean
Dim lPriceChange As PriceChange

For Each lSubscr In mDataSources
    If pTick.TickType = TickTypeMarketDepth Or pTick.TickType = TickTypeMarketDepthReset Then
        If lSubscr.SubscribedToMarketDepth Then fireRawMarketDepth lSubscr.DataSource, pTick
    ElseIf lSubscr.SubscribedToMarketData Then
        fireTick lSubscr.DataSource, pTick
        If lSubscr.DataSource.QuoteListeners.Count <> 0 Then
            If Not lGotQuote Then lQuote = getDataStream.GetCurrentQuote(pTick.TickType): lGotQuote = True
            fireQuote lSubscr.DataSource, lQuote, pTick.TickType
        End If
        If pTick.TickType = TickTypeTrade Or pTick.TickType = TickTypeClosePrice Then
            If lSubscr.DataSource.PriceChangeListeners.Count <> 0 Then
                If Not lGotPriceChange Then lPriceChange = getDataStream.GetPriceChange: lGotPriceChange = True
                firePriceChange lSubscr.DataSource, lPriceChange
            End If
        End If
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveDataSource"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Set lSubscr = getSubscription(pDataSource)

If lSubscr.SubscribedToMarketData Then
    lSubscr.SubscribedToMarketData = False
    mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
End If

If lSubscr.SubscribedToMarketData Then
    lSubscr.SubscribedToMarketData = False
    mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
End If

removeSubscription pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDataSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveMarketDataSubscription"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Set lSubscr = getSubscription(pDataSource)

If lSubscr.SubscribedToMarketData Then
    lSubscr.SubscribedToMarketData = False
    mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
End If

If Not lSubscr.SubscribedToMarketDepth Then removeSubscription pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDepthSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveMarketDepthSubscription"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Set lSubscr = getSubscription(pDataSource)

If lSubscr.SubscribedToMarketDepth Then
    lSubscr.SubscribedToMarketDepth = False
    mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions - 1
End If

If Not lSubscr.SubscribedToMarketData Then removeSubscription pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireError(ByVal pDataSource As MarketDataSource, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "fireError"
On Error GoTo Err

Dim lListener As ErrorListener
Dim ev As ErrorEventData

Set ev.Source = pDataSource
ev.ErrorCode = pErrorCode
ev.ErrorMessage = pErrorMessage

For Each lListener In pDataSource.ErrorListeners
    lListener.Notify ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePriceChange(ByVal pDataSource As MarketDataSource, ByRef pPriceChange As PriceChange)
Const ProcName As String = "firePriceChange"
On Error GoTo Err

Dim lListener As IPriceChangeListener
Dim ev As PriceChangeEventData

Set ev.Source = pDataSource
ev.PriceChange = pPriceChange

For Each lListener In pDataSource.PriceChangeListeners
    lListener.Change ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireQuote(ByVal pDataSource As MarketDataSource, ByRef pQuote As Quote, ByVal pTickType As TickTypes)
Const ProcName As String = "fireQuote"
On Error GoTo Err

Dim lListener As IQuoteListener
Dim ev As QuoteEventData

Set ev.Source = pDataSource
ev.Quote = pQuote

For Each lListener In pDataSource.QuoteListeners
    Select Case pTickType
    Case TickTypeBid
        lListener.Bid ev
    Case TickTypeAsk
        lListener.Ask ev
    Case TickTypeClosePrice
        lListener.PreviousClose ev
    Case TickTypeHighPrice
        lListener.High ev
    Case TickTypeLowPrice
        lListener.Low ev
    Case TickTypeMarketDepth
    Case TickTypeMarketDepthReset
    Case TickTypeTrade
        lListener.Trade ev
    Case TickTypeVolume
        lListener.Volume ev
    Case TickTypeOpenInterest
        lListener.OpenInterest ev
    Case TickTypeOpenPrice
        lListener.SessionOpen ev
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireRawMarketDepth( _
                ByVal pDataSource As MarketDataSource, _
                ByRef pTick As GenericTick)
Const ProcName As String = "fireRawMarketDepth"
On Error GoTo Err

Dim lListener As IRawMarketDepthListener
Dim ev As RawMarketDepthEventData

If pDataSource.RawMarketDepthListeners.Count = 0 Then Exit Sub

Set ev.Source = pDataSource

If pTick.TickType = TickTypeMarketDepth Then
    ev.MarketMaker = pTick.MarketMaker
    ev.Operation = pTick.Operation
    ev.Position = pTick.Position
    ev.Price = pTick.Price
    ev.Side = pTick.Side
    ev.Size = pTick.Size
End If

For Each lListener In pDataSource.RawMarketDepthListeners
    If pTick.TickType = TickTypeMarketDepth Then
        lListener.UpdateMarketDepth ev
    Else
        lListener.ResetMarketDepth ev
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireResetMarketDepth( _
                ByVal pDataSource As MarketDataSource)
Const ProcName As String = "fireResetMarketDepth"
On Error GoTo Err

If pDataSource.RawMarketDepthListeners.Count <> 0 Then
    Dim lRawListener As IRawMarketDepthListener
    Dim evRaw As RawMarketDepthEventData
    Set evRaw.Source = pDataSource
    For Each lRawListener In pDataSource.RawMarketDepthListeners
        lRawListener.ResetMarketDepth evRaw
    Next
ElseIf pDataSource.MarketDepthListeners.Count <> 0 Then
    Dim lListener As IMarketDepthListener
    Dim ev As MarketDepthEventData
    Set ev.Source = pDataSource
    For Each lListener In pDataSource.MarketDepthListeners
        lListener.ResetMarketDepth ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireSetMarketDepthCell( _
                ByVal pDataSource As MarketDataSource, _
                ByVal pSide As DOMSides, _
                ByVal pPrice As Double, _
                ByVal pSize As Long)
Const ProcName As String = "fireSetMarketDepthCell"
On Error GoTo Err

Dim lListener As IMarketDepthListener
Dim ev As MarketDepthEventData

If pDataSource.MarketDepthListeners.Count = 0 Then Exit Sub

Set ev.Source = pDataSource
ev.Price = pPrice
ev.Type = pSide
ev.Size = pSize
For Each lListener In pDataSource.MarketDepthListeners
    lListener.SetMarketDepthCell ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireTick(ByVal pDataSource As MarketDataSource, ByRef pTick As GenericTick)
Const ProcName As String = "fireTick"
On Error GoTo Err

Dim ev As GenericTickEventData
Dim lListener As IGenericTickListener

If pDataSource.GenericTickListeners.Count = 0 Then Exit Sub

Set ev.Source = pDataSource
ev.StreamId = pDataSource.StreamId
ev.Tick = pTick

For Each lListener In pDataSource.GenericTickListeners
    lListener.NotifyTick ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getDataStream() As MarketDataStream
Set getDataStream = mDataStreamRef.Target
End Function

Private Function getSubscription(ByVal pDataSource As MarketDataSource) As DataSourceSubscription
Const ProcName As String = "getSubscription"
On Error GoTo Err

Dim lKey As String
lKey = GetObjectKey(pDataSource)

If mDataSources.Contains(lKey) Then
    Set getSubscription = mDataSources.Item(lKey)
Else
    Set getSubscription = New DataSourceSubscription
    getSubscription.DataSource = pDataSource
    mDataSources.Add getSubscription, lKey
End If

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Function
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub removeSubscription(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "removeSubscription"
On Error GoTo Err

mDataSources.Remove GetObjectKey(pDataSource)

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub




