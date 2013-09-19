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
    If lSubscr.SubscribedToMarketDepth Then lSubscr.DataSource.FireResetMarketDepth
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
    If lSubscr.SubscribedToMarketDepth Then lSubscr.DataSource.FireSetMarketDepthCell pSide, pPrice, pSize
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Friend Sub Initialise(ByVal pDataStream As MarketDataStream)
Set mDataStreamRef = CreateWeakReference(pDataStream)
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
    If lSubscr.SubscribedToMarketData Then lSubscr.DataSource.FireError pErrorCode, pErrorMsg
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
    If lSubscr.SubscribedToMarketDepth Then lSubscr.DataSource.FireError pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

Dim lGotQuote As Boolean
Dim lGotPriceChange As Boolean

Dim lSubscr As DataSourceSubscription
For Each lSubscr In mDataSources
    Dim lDataSource As MarketDataSource
    Set lDataSource = lSubscr.DataSource
    
    If pTick.TickType = TickTypeMarketDepth Or pTick.TickType = TickTypeMarketDepthReset Then
        If lSubscr.SubscribedToMarketDepth Then lDataSource.FireRawMarketDepth pTick
    ElseIf lSubscr.SubscribedToMarketData Then
        lDataSource.FireTick pTick
        If lDataSource.HasQuoteListeners Then
            Dim lQuote As Quote
            If Not lGotQuote Then lQuote = getDataStream.GetCurrentQuote(pTick.TickType): lGotQuote = True
            lDataSource.FireQuote lQuote, pTick.TickType
        End If
        If pTick.TickType = TickTypeTrade Or pTick.TickType = TickTypeClosePrice Then
            If lDataSource.HasPriceChangeListeners Then
                Dim lPriceChange As PriceChange
                If Not lGotPriceChange Then lPriceChange = getDataStream.GetPriceChange: lGotPriceChange = True
                lDataSource.FirePriceChange lPriceChange
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




