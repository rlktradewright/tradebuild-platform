VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOMManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IDeferredAction

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Private Enum DeferredActions
    DeferredActionsNotifyCurrentMarketDepthValues
    DeferredActionsNotifyCurrentDom
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type PriceTableEntry
    Price                   As Double
    Size                    As Long
    MarketMaker             As String
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DOMManager"

'@================================================================================
' Member variables
'@================================================================================

Private mNumberOfMarketDepthRows                    As Long
Private mAskPrices()                                As PriceTableEntry
Private mMaxAskPricesIndex                          As Long
Private mBidPrices()                                As PriceTableEntry
Private mMaxBidPricesIndex                          As Long

Private mDataSources                                As DataSources

Private mPrevTradePrice                             As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData = Data

Select Case lData.Action
Case DeferredActionsNotifyCurrentMarketDepthValues
    doNotifyCurrentMarketDepthValues lData.Listener, lData.DataSource
Case DeferredActionsNotifyCurrentDom
    doNotifyCurrentDOM lData.Listener, lData.DataSource
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pNumberOfMarketDepthRows As Long, _
                ByVal pDataSources As DataSources)
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows
initialisePriceTables mNumberOfMarketDepthRows
Set mDataSources = pDataSources
End Sub

Friend Sub NotifyCurrentDOM( _
                ByVal pListener As IMarketDepthListener, _
                ByVal pDataSource As MarketDataSource)
Const ProcName As String = "NotifyCurrentDOM"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData.Action = DeferredActionsNotifyCurrentDom
Set lData.DataSource = pDataSource
Set lData.Listener = pListener

DeferAction Me, lData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCurrentMarketDepthValues( _
                ByVal pListener As IRawMarketDepthListener, _
                ByVal pDataSource As MarketDataSource)
Const ProcName As String = "NotifyCurrentMarketDepthValues"
On Error GoTo Err

Dim lData As DeferredNotificationActionData
lData.Action = DeferredActionsNotifyCurrentMarketDepthValues
Set lData.DataSource = pDataSource
Set lData.Listener = pListener

DeferAction Me, lData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessMarketDepth( _
                            pTick As GenericTick)
Const ProcName As String = "ProcessMarketDepth"
On Error GoTo Err

AssertArgument pTick.TickType = TickTypeMarketDepth, "Incorrect tick type"

If pTick.Side = DOMSides.DOMAsk Then
    updatePrices mAskPrices, _
                mMaxAskPricesIndex, _
                pTick.Side, _
                pTick.Position, _
                pTick.Operation, _
                pTick.Price, _
                pTick.Size
Else
    updatePrices mBidPrices, _
                mMaxBidPricesIndex, _
                pTick.Side, _
                pTick.Position, _
                pTick.Operation, _
                pTick.Price, _
                pTick.Size
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessMarketDepthReset()
Const ProcName As String = "ProcessMarketDepthReset"
On Error GoTo Err

initialisePriceTables mNumberOfMarketDepthRows
mPrevTradePrice = 0
mDataSources.DOMReset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ProcessTrade( _
                ByRef pTick As GenericTick)
              
Const ProcName As String = "ProcessTrade"
On Error GoTo Err

AssertArgument pTick.TickType = TickTypeTrade, "Incorrect tick type"

If pTick.Price <> mPrevTradePrice Then
    mDataSources.DOMSetCell DOMSides.DOMTrade, mPrevTradePrice, 0
End If

If pTick.Size <> 0 Then
    mDataSources.DOMSetCell DOMSides.DOMTrade, pTick.Price, pTick.Size
End If

mPrevTradePrice = pTick.Price

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Reset()
Const ProcName As String = "Reset"
On Error GoTo Err

initialisePriceTables mNumberOfMarketDepthRows
mPrevTradePrice = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Friend Sub doNotifyCurrentDOM( _
                ByVal pListener As IMarketDepthListener, _
                ByVal pDataSource As MarketDataSource)
Const ProcName As String = "doNotifyCurrentDOM"
On Error GoTo Err

If pDataSource.CurrentTick(TickTypeTrade).Price <> 0# Then
    Dim ev As MarketDepthEventData
    Set ev.Source = pDataSource
    ev.Price = pDataSource.CurrentTick(TickTypeTrade).Price
    ev.Size = pDataSource.CurrentTick(TickTypeTrade).Size
    ev.Type = DOMTrade
    pListener.SetMarketDepthCell ev
End If
notifyDOMSide DOMSides.DOMBid, mBidPrices, mMaxBidPricesIndex, pListener, pDataSource
notifyDOMSide DOMSides.DOMAsk, mAskPrices, mMaxAskPricesIndex, pListener, pDataSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub doNotifyCurrentMarketDepthValues( _
                ByVal pListener As IRawMarketDepthListener, _
                ByVal pDataSource As MarketDataSource)
Const ProcName As String = "doNotifyCurrentMarketDepthValues"
On Error GoTo Err

notifyMarketDepthValues mBidPrices, mMaxBidPricesIndex, pListener, pDataSource, DOMBid
notifyMarketDepthValues mAskPrices, mMaxAskPricesIndex, pListener, pDataSource, DOMAsk

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialisePriceTables(ByVal pNumberOfRows As Long)
ReDim mAskPrices(pNumberOfRows - 1) As PriceTableEntry
ReDim mBidPrices(pNumberOfRows - 1) As PriceTableEntry
mMaxAskPricesIndex = -1
mMaxBidPricesIndex = -1
End Sub

Private Sub notifyDOMSide( _
                ByVal pSide As DOMSides, _
                ByRef pPrices() As PriceTableEntry, _
                ByVal pNumEntries As Long, _
                ByVal pListener As IMarketDepthListener, _
                ByVal pDataSource As MarketDataSource)
Const ProcName As String = "notifyDOMSide"
On Error GoTo Err

Dim ev As MarketDepthEventData
Dim i As Long

Set ev.Source = pDataSource
ev.Type = pSide

For i = 0 To pNumEntries - 1
    If pPrices(i).Size <> 0 Then
        notifyDOMValue pPrices(i), pListener, ev
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyDOMValue( _
                ByRef pValue As PriceTableEntry, _
                ByVal pListener As IMarketDepthListener, _
                ByRef ev As MarketDepthEventData)
Const ProcName As String = "notifyDOMValue"
On Error GoTo Err

ev.Price = pValue.Price
ev.Size = pValue.Size
pListener.SetMarketDepthCell ev
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyMarketDepthValue( _
                ByRef pValue As PriceTableEntry, _
                ByVal pPosition As Long, _
                ByVal pOperation As DOMOperations, _
                ByVal pSide As DOMSides, _
                ByVal pListener As IRawMarketDepthListener, _
                ByRef ev As RawMarketDepthEventData)
Const ProcName As String = "notifyMarketDepthValue"
On Error GoTo Err

ev.MarketMaker = pValue.MarketMaker
ev.Operation = pOperation
ev.Position = pPosition
ev.Price = pValue.Price
ev.Side = pSide
ev.Size = pValue.Size
pListener.UpdateMarketDepth ev
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyMarketDepthValues( _
                ByRef pPrices() As PriceTableEntry, _
                ByVal pNumEntries As Long, _
                ByVal pListener As IRawMarketDepthListener, _
                ByVal pDataSource As MarketDataSource, _
                ByVal pSide As DOMSides)
Const ProcName As String = "notifyMarketDepthValues"
On Error GoTo Err

Dim ev As RawMarketDepthEventData
Dim i As Long
Dim lPosition As Long

Set ev.Source = pDataSource

lPosition = -1
For i = 0 To pNumEntries - 1
    If pPrices(i).Size <> 0 Then
        lPosition = lPosition + 1
        notifyMarketDepthValue pPrices(i), lPosition, DOMInsert, pSide, pListener, ev
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updatePrices(ByRef pPrices() As PriceTableEntry, _
                    ByRef pMaxIndex As Long, _
                    ByVal pSide As DOMSides, _
                    ByVal pPosition As Long, _
                    ByVal pOperation As DOMOperations, _
                    ByVal pPrice As Double, _
                    ByVal pSize As Long)
' NB: we don't update the screen if pPosition is 0, since these are
' passed earlier via the notifyBid and notifyAsk methods. Except for deletions
' (for example if the Bid drops, the current entry 0 must be deleted - though
' it could just be updated, TWS doesn't seem to work this way)

Const ProcName As String = "updatePrices"
On Error GoTo Err

Dim i As Long

Select Case pOperation
Case DOMOperations.DOMInsert
    If pPosition > pMaxIndex Then
        pMaxIndex = pPosition
    Else
        pMaxIndex = pMaxIndex + 1
    End If
Case DOMOperations.DOMUpdate
    If pPosition > pMaxIndex Then pMaxIndex = pPosition
Case DOMOperations.DOMDelete
    If pPosition > pMaxIndex Then
        pMaxIndex = pPosition
    End If
End Select

Do While pMaxIndex >= UBound(pPrices)
    ReDim Preserve pPrices(2 * (UBound(pPrices) + 1) - 1) As PriceTableEntry
Loop

Select Case pOperation
Case DOMOperations.DOMInsert
    For i = pMaxIndex - 1 To pPosition Step -1
        pPrices(i + 1) = pPrices(i)
    Next
    pPrices(pPosition).Price = pPrice
    pPrices(pPosition).Size = pSize
    mDataSources.DOMSetCell pSide, pPrice, pSize
Case DOMOperations.DOMUpdate
    If pPrice <> pPrices(pPosition).Price And pPrices(pPosition).Price <> 0 Then
        ' if need be we must blank the pSize display for the pPrice currently
        ' in this slot. Note that sometimes the market depth updates temporarily
        ' result in two or more adjacent slots for the same pPrice - therefore we
        ' only blank the pSize if the adjacent slot(s) are for a different pPrice.
        If pPosition = 0 Then
            If pPrices(1).Price <> pPrices(0).Price Then
                mDataSources.DOMSetCell pSide, pPrices(0).Price, 0
            End If
        Else
            If pPrices(pPosition + 1).Price <> pPrices(pPosition).Price And _
                pPrices(pPosition - 1).Price <> pPrices(pPosition).Price _
            Then
                mDataSources.DOMSetCell pSide, pPrices(pPosition).Price, 0
            End If
        End If
    End If
    pPrices(pPosition).Price = pPrice
    pPrices(pPosition).Size = pSize
    mDataSources.DOMSetCell pSide, pPrice, pSize
Case DOMOperations.DOMDelete
'    If (pPosition = 0 And pSide = DOMSides.DOMBid And pPrices(pPosition).Price = currentBid) _
'        Or _
'        (pPosition = 0 And pSide = DOMSides.DOMAsk And pPrices(pPosition).Price = currentAsk) _
'    Then
'        ' we get here when we've already set the Bid or Ask in the notifyBid/notifyAsk
'        ' methods, but the marketdepth updates are deleting some entries prior to
'        ' inserting the Bid/Ask (why it deletes an entry then immediately reinserts
'        ' it is a mystery to me, but it does!).
'        ' So in this scenario, we don't want to blank the pSize display for this pPrice
'    Else
'        ' otherwise we do want to blank the pSize entry for this pPrice
        If pPrices(pPosition).Price = 0 Then
            ' we getting a delete before a pPrice has been inserted in this pPosition,
            ' so just ignore it
        Else
            mDataSources.DOMSetCell pSide, pPrices(pPosition).Price, 0
        End If
'    End If
    
    For i = pPosition To pMaxIndex - 1
        pPrices(i) = pPrices(i + 1)
    Next
    pPrices(pMaxIndex).Price = 0
    pPrices(pMaxIndex).Size = 0
    pMaxIndex = pMaxIndex - 1
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




