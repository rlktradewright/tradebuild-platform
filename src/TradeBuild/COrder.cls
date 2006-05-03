VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order"
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
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

Event StatusChanged()
Event Fill( _
                ByVal quantityFilled As Long, _
                ByVal fillPrice As Double)

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

Private mTicker         As ticker
Private mContract       As TradeBuild.Contract

Private mDummy          As Boolean

Private mDifferent      As Boolean
' The following fields are values received from TWS in the openOrder2 -
' they are only set when there is a discrepancy between TWS's info
' and ours
Private mTWSAction As TradeBuild.OrderActions
Private mTWSQuantity As Long
Private mTWSOrderType As TradeBuild.OrderTypes
Private mTWSLimitPrice As Double
Private mTWSTriggerPrice As Double

Private mFillTime       As Date

Private mMarketIndex    As Long
Private mAtExchange     As Boolean
Private mSimulated      As Boolean
Private mStatus         As OrderStatuses
Private mIdentifier     As String
Private mParentId       As String

Private mBrokerId       As String
Private mClientID       As Long
'Private mPermId         As Long
    
Private mAction         As OrderActions
Private mQuantity       As Long
Private mQuantityFilled As Long
Private mAveragePrice   As Double
Private mLastFillPrice  As Double
Private mOrderType      As TradeBuild.OrderTypes
Private mLimitPrice     As Double
Private mTriggerPrice   As Double
Private mOffset         As Long

    ' extended order fields
Private mTif            As OrderTifs
Private mOcaGroup       As String
Private mAccount        As String
Private mOpenClose      As String
Private mOrigin         As OrderOrigins
Private mOrderRef       As String
Private mTransmit       As Boolean
Private mBlockOrder     As Boolean
Private mSweepToFill    As Boolean
Private mDisplaySize    As Long
Private mTriggerMethod  As TriggerMethods
Private mIgnoreRegularTradingHours      As Boolean
Private mHidden         As Boolean
Private mDiscretionaryAmount As Double
Private mGoodAfterTime  As Date
Private mGoodTillDate   As Date
    
    ' share allocation
Private mSharesAllocation As String
    
    'fa
Private mFAGroup        As String
Private mFAProfile      As String
Private mFAMethod       As String
Private mFAPercentage   As String
    
    
'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mStatus = OrderStatuses.OrderStatusCreated
mAction = ActionBuy
mOrderType = OrderTypeNone
mTif = TIFDay
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let account(ByVal value As String)
mAccount = value
End Property

Public Property Get account() As String
account = mAccount
End Property

Public Property Let action(ByVal value As OrderActions)
mAction = UCase$(value)
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Friend Property Let atExchange(ByVal value As Boolean)
mAtExchange = value
End Property

Friend Property Get atExchange() As Boolean
atExchange = mAtExchange
End Property

Public Property Let averagePrice(ByVal value As Double)
mAveragePrice = value
End Property

Public Property Get averagePrice() As Double
averagePrice = mAveragePrice
End Property

Public Property Let blockOrder(value As Boolean)
mBlockOrder = value
End Property

Public Property Get blockOrder() As Boolean
blockOrder = mBlockOrder
End Property

Friend Property Let brokerId(ByVal value As String)
mBrokerId = value
End Property

Public Property Get brokerId() As String
brokerId = mBrokerId
End Property

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Public Property Get clientID() As Long
clientID = mClientID
End Property

Public Property Get Contract() As TradeBuild.Contract
Set Contract = mContract
End Property

Public Property Let displaySize(ByVal value As Long)
mDisplaySize = value
End Property

Public Property Get displaySize() As Long
displaySize = mDisplaySize
End Property

Public Property Let discretionaryAmount(ByVal value As Double)
mDiscretionaryAmount = value
End Property

Public Property Get discretionaryAmount() As Double
discretionaryAmount = mDiscretionaryAmount
End Property

Friend Property Let dummy(ByVal value As Boolean)
mDummy = value
End Property

Public Property Get dummy() As Boolean
dummy = mDummy
End Property

Public Property Let FAGroup(ByVal value As String)
mFAGroup = value
End Property

Public Property Get FAGroup() As String
FAGroup = mFAGroup
End Property

Public Property Let FAMethod(ByVal value As String)
mFAMethod = value
End Property

Public Property Get FAMethod() As String
FAMethod = mFAMethod
End Property

Public Property Let FAPercentage(ByVal value As String)
mFAPercentage = value
End Property

Public Property Get FAPercentage() As String
FAPercentage = mFAPercentage
End Property

Public Property Let FAProfile(ByVal value As String)
mFAProfile = value
End Property

Public Property Get FAProfile() As String
FAProfile = mFAProfile
End Property

Public Property Let fillTime(ByVal value As Date)
mFillTime = value
End Property

Public Property Get fillTime() As Date
fillTime = mFillTime
End Property

Public Property Let goodAfterTime(ByVal value As Date)
mGoodAfterTime = value
End Property

Public Property Get goodAfterTime() As Date
goodAfterTime = mGoodAfterTime
End Property

Public Property Let goodTillDate(ByVal value As Date)
mGoodTillDate = value
End Property

Public Property Get goodTillDate() As Date
goodTillDate = mGoodTillDate
End Property

Public Property Let hidden(value As Boolean)
mHidden = value
End Property

Public Property Get hidden() As Boolean
hidden = mHidden
End Property

Friend Property Let id(ByVal value As String)
mIdentifier = value
End Property

Public Property Get id() As String
id = mIdentifier
End Property

Public Property Let ignoreRegularTradingHours(value As Boolean)
mIgnoreRegularTradingHours = value
End Property

Public Property Get ignoreRegularTradingHours() As Boolean
ignoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Friend Property Let lastFillPrice(ByVal value As Double)
mLastFillPrice = value
End Property

Public Property Get lastFillPrice() As Double
lastFillPrice = mLastFillPrice
End Property

Public Property Let limitPrice(ByVal value As Double)
mLimitPrice = value
End Property

Public Property Get limitPrice() As Double
limitPrice = mLimitPrice
End Property

Friend Property Let marketIndex(ByVal value As Long)
mMarketIndex = value
End Property

Friend Property Get marketIndex() As Long
marketIndex = mMarketIndex
End Property

Public Property Let ocaGroup(ByVal value As String)
mOcaGroup = value
End Property

Public Property Get ocaGroup() As String
ocaGroup = mOcaGroup
End Property

Friend Property Let offset(ByVal value As Long)
mOffset = value
End Property

Friend Property Get offset() As Long
offset = mOffset
End Property

Public Property Let openClose(ByVal value As String)
mOpenClose = value
End Property

Public Property Get openClose() As String
openClose = mOpenClose
End Property

Public Property Let orderRef(ByVal value As String)
mOrderRef = value
End Property

Public Property Get orderRef() As String
orderRef = mOrderRef
End Property

Public Property Let orderType(ByVal value As TradeBuild.OrderTypes)
Select Case value
Case OrderTypeNone
Case OrderTypeMarket
Case OrderTypeMarketOnClose
Case OrderTypeLimit
Case OrderTypeLimitOnClose
Case OrderTypePeggedToMarket
Case OrderTypeStop
Case OrderTypeStopLimit
Case OrderTypeTrail
Case OrderTypeRelative
Case OrderTypeVWAP
Case OrderTypeMarketToLimit
Case OrderTypeQuote
Case OrderTypeAutoStop
Case OrderTypeAutoLimit
Case OrderTypeAdjust
Case OrderTypeAlert
Case OrderTypeLimitIfTouched
Case OrderTypeMarketIfTouched
Case OrderTypeTrailLimit
Case OrderTypeMarketWithProtection
Case OrderTypeMarketOnOpen
Case OrderTypeLimitOnOpen
Case OrderTypePeggedToPrimary
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
            "TradeBuild.Order::(let)orderType", _
            "Invalid order type " & value
End Select
mOrderType = value
End Property

Public Property Get orderType() As TradeBuild.OrderTypes
orderType = mOrderType
End Property

Public Property Let origin(ByVal value As OrderOrigins)
mOrigin = value
End Property

Public Property Get origin() As OrderOrigins
origin = mOrigin
End Property

Friend Property Let parentId(ByVal value As String)
mParentId = value
End Property

Friend Property Get parentId() As String
parentId = mParentId
End Property

'Public Property Let permId(value As Long)
'mPermId = value
'End Property
'
'Public Property Get permId() As Long
'permId = mPermId
'End Property

Public Property Let quantity(ByVal value As Long)
mQuantity = value
End Property

Public Property Get quantity() As Long
quantity = mQuantity
End Property

Public Property Let quantityFilled(ByVal value As Long)
mQuantityFilled = value
End Property

Public Property Get quantityFilled() As Long
quantityFilled = mQuantityFilled
End Property

Public Property Let sharesAllocation(value As String)
mSharesAllocation = value
End Property

Public Property Get sharesAllocation() As String
sharesAllocation = mSharesAllocation
End Property

Public Property Let simulated(ByVal value As Boolean)
mSimulated = value
End Property

Public Property Get simulated() As Boolean
simulated = mSimulated
End Property

Public Property Let status(value As OrderStatuses)
mStatus = value
End Property

Public Property Get status() As OrderStatuses
status = mStatus
End Property

Public Property Let sweepToFill(value As Boolean)
mSweepToFill = value
End Property

Public Property Get sweepToFill() As Boolean
sweepToFill = mSweepToFill
End Property

Friend Property Let ticker(ByVal value As ticker)
Set mTicker = value
Set mContract = mTicker.Contract
End Property

Public Property Get ticker() As ticker
Set ticker = mTicker
End Property

Public Property Let timeInForce(ByVal value As OrderTifs)
mTif = value
End Property

Public Property Get timeInForce() As OrderTifs
timeInForce = mTif
End Property

Public Property Let transmit(value As Boolean)
mTransmit = value
End Property

Public Property Get transmit() As Boolean
transmit = mTransmit
End Property

Public Property Let triggerMethod(value As TriggerMethods)
mTriggerMethod = value
End Property

Public Property Get triggerMethod() As TriggerMethods
triggerMethod = mTriggerMethod
End Property

Public Property Let triggerPrice(ByVal value As Double)
mTriggerPrice = value
End Property

Public Property Get triggerPrice() As Double
triggerPrice = mTriggerPrice
End Property

'================================================================================
' Methods
'================================================================================

Public Function clone() As order
Set clone = New order
With clone
    .ticker = mTicker
    .marketIndex = mMarketIndex
    .simulated = mSimulated
    .status = mStatus
    .id = mIdentifier
    .clientID = mClientID
    .brokerId = mBrokerId
    .action = mAction
    .quantity = mQuantity
    .quantityFilled = mQuantityFilled
    .averagePrice = mAveragePrice
    .lastFillPrice = mLastFillPrice
    .orderType = mOrderType
    .limitPrice = mLimitPrice
    .triggerPrice = mTriggerPrice
    .timeInForce = mTif
    .ocaGroup = mOcaGroup
    .account = mAccount
    .openClose = mOpenClose
    .origin = mOrigin
    .orderRef = mOrderRef
    .transmit = mTransmit
    .parentId = mParentId
    .blockOrder = mBlockOrder
    .sweepToFill = mSweepToFill
    .displaySize = mDisplaySize
    .triggerMethod = mTriggerMethod
    .ignoreRegularTradingHours = mIgnoreRegularTradingHours
    .hidden = mHidden
    .discretionaryAmount = mDiscretionaryAmount
    .goodAfterTime = mGoodAfterTime
    .sharesAllocation = mSharesAllocation
    .FAGroup = mFAGroup
    .FAProfile = mFAProfile
    .FAMethod = mFAMethod
    .FAPercentage = mFAPercentage
End With
End Function

Friend Sub Fill(ByVal quantity As Long, ByVal price As Double)
' only used by the Order Simulator - when the Order Simulator is changed to be a
' service provider, this should be removed
mLastFillPrice = price
mAveragePrice = (mQuantityFilled * mAveragePrice + quantity * price) / (mQuantityFilled + quantity)
mQuantityFilled = mQuantityFilled + quantity
mQuantity = mQuantity - quantity
End Sub

Friend Sub finish()
Set mTicker = Nothing
Set mContract = Nothing
End Sub

Public Function isAttributeModifiable( _
                ByVal attributeID As OrderAttributeIds) As Boolean
Select Case attributeID
Case OrderAttributeIds.action
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.blockOrder
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.displaySize
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.discretionaryAmount
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.goodAfterTime
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.goodTillDate
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.hidden
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.ignoreRegularTradingHours
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.limitPrice
    Select Case mOrderType
    Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeLimitOnOpen
    Case Else
        Exit Function
    End Select
    
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted, _
        OrderStatusSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.orderType
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.quantity
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted, _
        OrderStatusSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.sweepToFill
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.timeInForce
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case OrderAttributeIds.triggerPrice
    Select Case mOrderType
    Case OrderTypeStop, _
        OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeMarketIfTouched
    Case Else
        Exit Function
    End Select
    
    Select Case mStatus
    Case OrderStatusCreated, _
        OrderStatusRejected, _
        OrderStatusPendingSubmit, _
        OrderStatusPreSubmitted
        isAttributeModifiable = True
    Case Else
    End Select
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
            "TradeBuild.Order::isAttributeModifiable", _
            "Invalid attribute id: " & attributeID
End Select
End Function

Public Function isModifiable() As Boolean
If mStatus <> OrderStatuses.OrderStatusCancelled And _
    mStatus <> OrderStatuses.OrderStatusCancelling And _
    mStatus <> OrderStatuses.OrderStatusFilled _
Then isModifiable = True
End Function

Friend Sub openOrder(ByVal action As TradeBuild.OrderActions, _
                        ByVal quantity As Long, _
                        ByVal orderType As TradeBuild.OrderTypes, _
                        ByVal lmtPrice As Double, _
                        ByVal auxPrice As Double, _
                        ByVal tif As TradeBuild.OrderTifs, _
                        ByVal ocaGroup As String)
                        
mTWSAction = action
If action <> mAction Then
    mDifferent = True
End If

mTWSQuantity = quantity
If quantity <> mQuantity Then
    mDifferent = True
End If

mTWSOrderType = orderType
If orderType <> mOrderType Then
    mDifferent = True
End If
If orderType = TradeBuild.OrderTypes.OrderTypeLimit Or _
    orderType = TradeBuild.OrderTypes.OrderTypeStopLimit _
Then
    mTWSLimitPrice = lmtPrice
    If lmtPrice <> mLimitPrice Then
        mDifferent = True
    End If
Else
    mTWSLimitPrice = 0
End If
If orderType = TradeBuild.OrderTypes.OrderTypeStop Or _
    orderType = TradeBuild.OrderTypes.OrderTypeStopLimit _
Then
    mTWSTriggerPrice = auxPrice
    If auxPrice <> mTriggerPrice Then
        mDifferent = True
    End If
Else
    mTWSTriggerPrice = 0
End If
' NB: mDifferent will be cleared when updateStatus is called (this always
' happens shortly after openOrder)
'If mDifferent Then Stop
End Sub

Friend Sub updateStatus(ByVal pStatus As TradeBuild.OrderStatuses, _
                ByVal filled As Long, _
                ByVal remaining As Long, _
                ByVal avgFillPrice As Double, _
                ByVal lastFillPrice As Double, _
                ByVal timestamp As Date)
Dim filledThisTime As Long

If pStatus = TradeBuild.OrderStatuses.OrderStatusPendingSubmit Then
    mStatus = pStatus
    RaiseEvent StatusChanged
    Exit Sub
End If

If filled > mQuantityFilled Then
    filledThisTime = filled - mQuantityFilled
    If Not mDummy Then mTicker.notifyListeners _
                            mContract.specifier.exchange & "," & _
                            mContract.specifier.symbol & "," & _
                            mContract.specifier.expiry & "," & _
                            filledThisTime & "," & _
                            lastFillPrice & "," & _
                            orderActionToString(mAction) & "," & _
                            timestamp & "," & _
                            id & "," & _
                            brokerId, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTOrderDetailDummy, TradeBuildListenValueTypes.VTOrderDetail)
End If

mQuantityFilled = filled
mQuantity = mQuantity - mQuantityFilled
If mQuantity <> remaining And _
    pStatus <> TradeBuild.OrderStatuses.OrderStatusCancelled _
Then
    ' TWS must have a different view of the quantity from us. This could
    ' be because we just changed the quantity but a fill occurred just before
    ' TWS actioned the change
    mDifferent = True
    mTWSQuantity = remaining
End If

mAveragePrice = avgFillPrice
mLastFillPrice = lastFillPrice

If mStatus = TradeBuild.OrderStatuses.OrderStatusCancelling Then
    If pStatus = TradeBuild.OrderStatuses.OrderStatusSubmitted Or _
        pStatus = TradeBuild.OrderStatuses.OrderStatusPreSubmitted _
    Then
        ' this could mean:
        '   EITHER  we tried to cancel the order but TWS had recovered from
        '           a loss of connection to the IB server and at the time
        '           of cancelling we hadn't reassociated order ids to
        '           the client. When we did reassociate, it caused this
        '           status update notification. So we now need to do the
        '           cancel again.
        '   OR      we tried to cancel the order but got a partial fill
        '           before TWS could action the cancel. We should get
        '           a 'canceled' status update notification shortly, but
        '           since we can't distinguish this case from the other we
        '           just submit the cancel again.
        mTicker.notifyListeners _
                    "Re-cancelling order " & mIdentifier & " - discrepancies with server:" & vbCrLf & _
                    "    Me:  status: " & orderStatusToString(mStatus) & vbCrLf & _
                    "    TWS: status: " & orderStatusToString(status), _
                    TradeBuildListenValueTypes.VTTWS
        ' before doing the cancel we need to put the order back into
        ' the simulator
        Stop
        'gTradeBuildAPI.cancelOrder mIdentifier
        Exit Sub
    End If
End If
mStatus = pStatus


If filledThisTime <> 0 Then
    mFillTime = timestamp
    RaiseEvent Fill(filledThisTime, lastFillPrice)
End If
RaiseEvent StatusChanged

If mDifferent Then
    If mStatus <> TradeBuild.OrderStatuses.OrderStatusFilled Then
        ' must re-place the order
        mTicker.notifyListeners _
                    "Resubmitting order " & mIdentifier & " - discrepancies with TWS:" & vbCrLf & _
                    "    Me:" & vbCrLf & _
                    IIf(mTWSAction <> 0, "         action:        " & orderActionToString(mAction) & vbCrLf, "") & _
                    IIf(mTWSQuantity <> 0, "         quantity:      " & mQuantity & vbCrLf, "") & _
                    IIf(mTWSOrderType <> 0, "         order type:    " & orderTypeToString(mOrderType) & vbCrLf, "") & _
                    IIf(mTWSLimitPrice <> 0, "         limit price:   " & mLimitPrice & vbCrLf, "") & _
                    IIf(mTWSTriggerPrice <> 0, "         trigger price: " & mTriggerPrice & vbCrLf, "") & _
                    "    TWS:" & vbCrLf & _
                    IIf(mTWSAction <> 0, "         action:        " & orderActionToString(mTWSAction) & vbCrLf, "") & _
                    IIf(mTWSQuantity <> 0, "         quantity:      " & mTWSQuantity & vbCrLf, "") & _
                    IIf(mTWSOrderType <> 0, "         order type:    " & orderTypeToString(mTWSOrderType) & vbCrLf, "") & _
                    IIf(mTWSLimitPrice <> 0, "         limit price:   " & mTWSLimitPrice & vbCrLf, "") & _
                    IIf(mTWSTriggerPrice <> 0, "         trigger price: " & mTWSTriggerPrice & vbCrLf, ""), _
                    TradeBuildListenValueTypes.VTTWS
        mTransmit = True
        Stop
        'gTradeBuildAPI.placeOrder Me
    End If
    mDifferent = False
End If

End Sub

'================================================================================
' Helper Functions
'================================================================================








