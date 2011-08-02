VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event Ready()
Event NotReady()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "OrderContext"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName                           As String

Private mDataSource                     As IMarketDataSource
Attribute mDataSource.VB_VarHelpID = -1
Private mOrderSubmitter                 As IOrderSubmitter
Private mOrderAuthoriser                As IOrderAuthoriser
Private mAccumulatedBracketOrders       As Orders
Private mAccumulatedOrders              As Orders


Private mIsSimulated                    As Boolean

Private mActiveBracketOrders            As BracketOrders
    ' this contains all bracket orders for which one or more orders are
    ' not yet complete (or not yet cancelled)
    
Private mClosedBracketOrders            As BracketOrders
    ' this contains all bracket orders for which all orders have completed (or
    ' been cancelled)
    
Private mCompletedSize                  As Long
    ' the total Size for all completed bracket orders (may be non-zero if one
    ' or more completed bracket orders has no target or stop orders)
    
Private mPermittedOrderAttributes       As Long
Private mPermittedOrderTifs             As Long
Private mPermittedOrderTypes            As Long
Private mPermittedStopTriggerMethods    As Long

Private mEnabled                        As Boolean
Private mIsReady                        As Boolean

Private mTemplateOrder                  As Order

Private mPreventUnprotectedPositions    As Boolean

'Private mOrigin As OrderOrigins
'Private mBlockOrder As Boolean
'Private mSweepToFill As Boolean
'Private mDisplaySize As Long
'Private mIgnoreRegularTradingHours As Boolean
'Private mHidden As Boolean
'Private mDiscretionaryAmount As Double
'Private mGoodAfterTime As Date
'Private mGoodAfterTimeTZ    As String
'Private mGoodTillDate   As Date
'Private mGoodTillDateTZ As String
''Private mRule80A               ' not implemented
'Private mSettlingFirm As String
'Private mAllOrNone As Boolean
'Private mMinimumQuantity As Long
'Private mPercentOffset As Double
'Private mETradeOnly As Boolean
'Private mFirmQuoteOnly  As Boolean
'Private mNBBOPriceCap As Double
'Private mOverrideConstraints  As Boolean
'Private mLimitPrice As Double
'Private mOrderType As OrderTypes
'Private mQuantity As Long
'Private mTimeInForce As OrderTIFs
'Private mTriggerPrice As Double
'Private mStopTriggerMethod As StopTriggerMethods

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mActiveBracketOrders = New BracketOrders
Set mClosedBracketOrders = New BracketOrders
Set mTemplateOrder = New Order
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderContext terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ActiveBracketOrders() As BracketOrders
Set ActiveBracketOrders = mActiveBracketOrders
End Property

Public Property Get ActiveSize() As Long
Dim lBracketOrder As BracketOrder

Const ProcName As String = "ActiveSize"

On Error GoTo Err

For Each lBracketOrder In mActiveBracketOrders
    ActiveSize = ActiveSize + lBracketOrder.size
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'Public Property Let AllOrNone(ByVal value As Boolean)
'Const ProcName As String = "AllOrNone"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttAllOrNone
'mAllOrNone = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get AllOrNone() As Boolean
'AllOrNone = mAllOrNone
'End Property
'
'Public Property Let BlockOrder(ByVal value As Boolean)
'Const ProcName As String = "BlockOrder"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttBlockOrder
'mBlockOrder = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get BlockOrder() As Boolean
'BlockOrder = mBlockOrder
'End Property
'
Public Property Get ClosedBracketOrders() As BracketOrders
Set ClosedBracketOrders = mClosedBracketOrders
End Property

Public Property Get Complete() As Boolean
Dim op As BracketOrder
Const ProcName As String = "Complete"

On Error GoTo Err

For Each op In mActiveBracketOrders
    If Not op.IsSimulated Then Complete = False: Exit Property
Next
Complete = True

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CompletedSize() As Long
CompletedSize = mCompletedSize
End Property

Public Property Get DataSource() As IMarketDataSource
Set Ticker = mDataSource
End Property

'Public Property Let DisplaySize(ByVal value As Long)
'Const ProcName As String = "DisplaySize"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttDisplaySize
'mDisplaySize = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get DisplaySize() As Long
'DisplaySize = mDisplaySize
'End Property

Friend Property Let Enabled(ByVal value As Boolean)
Const ProcName As String = "Enabled"

On Error GoTo Err

mEnabled = value
checkReady

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'Public Property Let ETradeOnly(ByVal value As Boolean)
'Const ProcName As String = "ETradeOnly"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttETradeOnly
'mETradeOnly = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get ETradeOnly() As Boolean
'ETradeOnly = mETradeOnly
'End Property
'
'Public Property Let FirmQuoteOnly(ByVal value As Boolean)
'Const ProcName As String = "FirmQuoteOnly"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttFirmQuoteOnly
'mFirmQuoteOnly = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get FirmQuoteOnly() As Boolean
'FirmQuoteOnly = mFirmQuoteOnly
'End Property
'
'Public Property Let GoodAfterTime(ByVal value As Date)
'Const ProcName As String = "GoodAfterTime"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttGoodAfterTime
'mGoodAfterTime = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get GoodAfterTime() As Date
'GoodAfterTime = mGoodAfterTime
'End Property
'
'Public Property Let GoodAfterTimeTZ(ByVal value As String)
'Const ProcName As String = "GoodAfterTimeTZ"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttGoodAfterTimeTZ
'mGoodAfterTimeTZ = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get GoodAfterTimeTZ() As String
'GoodAfterTimeTZ = mGoodAfterTimeTZ
'End Property
'
'Public Property Let GoodTillDate(ByVal value As Date)
'Const ProcName As String = "GoodTillDate"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttGoodTillDate
'mGoodTillDate = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get GoodTillDate() As Date
'GoodTillDate = mGoodTillDate
'End Property
'
'Public Property Let GoodTillDateTZ(ByVal value As String)
'Const ProcName As String = "GoodTillDateTZ"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttGoodTillDateTZ
'mGoodTillDateTZ = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get GoodTillDateTZ() As String
'GoodTillDateTZ = mGoodTillDateTZ
'End Property
'
'Public Property Let Hidden(ByVal value As Boolean)
'Const ProcName As String = "Hidden"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttHidden
'mHidden = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get Hidden() As Boolean
'Hidden = mHidden
'End Property
'
'Public Property Let IgnoreRegularTradingHours(ByVal value As Boolean)
'Const ProcName As String = "IgnoreRegularTradingHours"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttIgnoreRTH
'mIgnoreRegularTradingHours = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get IgnoreRegularTradingHours() As Boolean
'IgnoreRegularTradingHours = mIgnoreRegularTradingHours
'End Property

Friend Property Let IsReady(ByVal value As Boolean)
Const ProcName As String = "IsReady"

On Error GoTo Err

mIsReady = value
checkReady

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsReady() As Boolean
IsReady = (mIsReady And mEnabled)
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

'Public Property Let LimitPrice(ByVal value As Double)
'Const ProcName As String = "LimitPrice"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttLimitPrice
'mLimitPrice = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get LimitPrice() As Double
'LimitPrice = mLimitPrice
'End Property
'
'Public Property Let MinimumQuantity(ByVal value As Long)
'Const ProcName As String = "MinimumQuantity"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttMinimumQuantity
'mMinimumQuantity = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get MinimumQuantity() As Long
'MinimumQuantity = mMinimumQuantity
'End Property

Friend Property Get Name() As String
Name = mName
End Property

'Public Property Let NbboPriceCap(ByVal value As Double)
'Const ProcName As String = "NbboPriceCap"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttNBBOPriceCap
'mNBBOPriceCap = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get NbboPriceCap() As Double
'NbboPriceCap = mNBBOPriceCap
'End Property
'
'Public Property Let OrderType(ByVal value As OrderTypes)
'Const ProcName As String = "OrderType"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttOrderType
'If Not IsOrderTypeValueSupported(value) Then attributeValueNotSupported OrderAttributes.OrderAttOrderType, value
'mOrderType = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get OrderType() As OrderTypes
'OrderType = mOrderType
'End Property
'
'Public Property Let Origin(ByVal value As OrderOrigins)
'Const ProcName As String = "Origin"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttOrigin
'mOrigin = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get Origin() As OrderOrigins
'Origin = mOrigin
'End Property

'Public Property Let OverrideConstraints(ByVal value As Boolean)
'Const ProcName As String = "OverrideConstraints"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttOverrideConstraints
'mOverrideConstraints = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get OverrideConstraints() As Boolean
'OverrideConstraints = mOverrideConstraints
'End Property

Public Property Get PendingSize() As Long
Dim lBracketOrder As BracketOrder
Const ProcName As String = "PendingSize"

On Error GoTo Err

For Each lBracketOrder In mActiveBracketOrders
    PendingSize = PendingSize + lBracketOrder.PendingSize
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'Public Property Let PercentOffset(ByVal value As Double)
'Const ProcName As String = "PercentOffset"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttPercentOffset
'mPercentOffset = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get PercentOffset() As Double
'PercentOffset = mPercentOffset
'End Property

Public Property Get PermittedOrderTifs() As Long
PermittedOrderTifs = mPermittedOrderTifs
End Property

Public Property Get PermittedOrderTypes() As Long
PermittedOrderTypes = mPermittedOrderTypes
End Property

Public Property Get PermittedStopTriggerMethods() As Long
PermittedStopTriggerMethods = mPermittedStopTriggerMethods
End Property

Public Property Let PreventUnprotectedPositions( _
                ByVal value As Boolean)
mPreventUnprotectedPositions = value
End Property

Public Property Get PreventUnprotectedPositions() As Boolean
PreventUnprotectedPositions = mPreventUnprotectedPositions
End Property

'Public Property Let Quantity(ByVal value As Long)
'Const ProcName As String = "Quantity"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttQuantity
'mQuantity = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get Quantity() As Long
'Quantity = mQuantity
'End Property

Public Property Get Risk() As Currency
Dim lBracketOrder As BracketOrder
Const ProcName As String = "Risk"

On Error GoTo Err

For Each lBracketOrder In mActiveBracketOrders
    If lBracketOrder.Risk = MaxCurrency Then
        Risk = MaxCurrency
        Exit For
    End If
    Risk = Risk + lBracketOrder.Risk
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'Public Property Let SettlingFirm(ByVal value As String)
'Const ProcName As String = "SettlingFirm"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttSettlingFirm
'mSettlingFirm = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get SettlingFirm() As String
'SettlingFirm = mSettlingFirm
'End Property

Public Property Get size() As Long
Dim lBracketOrder As BracketOrder

Const ProcName As String = "Size"

On Error GoTo Err

size = mCompletedSize

For Each lBracketOrder In mActiveBracketOrders
    size = size + lBracketOrder.size
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'Public Property Let StopTriggerMethod(ByVal value As StopTriggerMethods)
'Const ProcName As String = "StopTriggerMethod"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttStopTriggerMethod
'If Not IsStopTriggerMethodValueSupported(value) Then attributeValueNotSupported OrderAttributes.OrderAttStopTriggerMethod, value
'mStopTriggerMethod = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get StopTriggerMethod() As StopTriggerMethods
'StopTriggerMethod = mStopTriggerMethod
'End Property
'
'Public Property Let SweepToFill(ByVal value As Boolean)
'Const ProcName As String = "SweepToFill"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttSweepToFill
'mSweepToFill = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get SweepToFill() As Boolean
'SweepToFill = mSweepToFill
'End Property

Public Property Get TemplateOrder() As Order
Set TemplateOrder = mTemplateOrder
End Property

'Public Property Let TimeInForce(ByVal value As OrderTIFs)
'Const ProcName As String = "TimeInForce"
'
'On Error GoTo Err
'
'If Not gIsValidTIF(value) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be a member of the OrderTifs enum"
'
'checkAttributeSupported OrderAttributes.OrderAttTimeInForce
'If Not IsOrderTifValueSupported(value) Then attributeValueNotSupported OrderAttributes.OrderAttTimeInForce, value
'mTimeInForce = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get TimeInForce() As OrderTIFs
'TimeInForce = mTimeInForce
'End Property
'
'Public Property Let TriggerPrice(ByVal value As Double)
'Const ProcName As String = "TriggerPrice"
'
'On Error GoTo Err
'
'checkAttributeSupported OrderAttributes.OrderAttTriggerPrice
'mTriggerPrice = value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Property
'
'Public Property Get TriggerPrice() As Double
'TriggerPrice = mTriggerPrice
'End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ActivateBracketOrder( _
                ByVal pBracketOrder As BracketOrder)
Const ProcName As String = "ActivateBracketOrder"

On Error GoTo Err

pBracketOrder.Ticker = mDataSource
pBracketOrder.OrderSubmitter = mOrderSubmitter
If pBracketOrder.IsComplete Then
    NotifyBracketOrderClosed pBracketOrder
Else
    mActiveBracketOrders.Add pBracketOrder
End If
pBracketOrder.Activate

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

''
' Creates a buy<code>BracketOrder</code>
'
' @return the <code>BracketOrder</code> object created
'
' @param number         quantity to buy
'       EntryOrderType  the type of order to enter the position:
'           EntryOrderTypeMarket            a market order
'           EntryOrderTypeMarketOnOpen      a market on open order
'           EntryOrderTypeMarketOnClose     a market on close order
'           EntryOrderTypeMarketIfTouched   a market if touched order
'           EntryOrderTypeMarketToLimit     a market to limit order
'           EntryOrderTypeBid               a limit order at the current Bid
'           EntryOrderTypeAsk               a limit order at the current Ask
'           EntryOrderTypeLast              a limit order at the last Trade price
'           EntryOrderTypeLimit             a limit order
'           EntryOrderTypeLimitOnOpen       a limit on open order
'           EntryOrderTypeLimitOnClose      a limit on close order
'           EntryOrderTypeLimitIfTouched    a limit if touched order
'           EntryOrderTypeStop              a stop order
'           EntryOrderTypeStopLimit         a stop limit order
' @param BuyPrice       the limit price for a limit order (ignored for EntryOrderTypeBid,
'                       EntryOrderTypeAsk and EntryOrderTypeLast)
' @param BuyOffset      number of ticks to modify the BuyPrice. For example, to place an order
'                       one tick below the current Ask, use EntryOrderTypeAsk and set
'                       BuyOffset to -1
' @param TriggerPrice   used with EntryOrderTypeMarketIfTouched, EntryOrderTypeLimitIfTouched,
'                       EntryOrderTypeStop and EntryOrderTypeStopLimit to specify the price
'                       at which the order is triggered
' @param StopLossOrderType  the type of stop loss order to place:
'           StopLossOrderTypeNone               no stop loss order
'           StopLossOrderTypeStop               use a stop order
'           StopLossOrderTypeStopLimit          use a stop limit order
'           StopLossOrderTypeBid                use a stop order at the current Bid (see StopOffset)
'           StopLossOrderTypeAsk                use a stop order at the current Ask (see StopOffset)
'           StopLossOrderTypeLast               use a stop order at the last Trade price (see StopOffset)
'           StopLossOrderTypeAuto               use a stop order at the fill price for the
'                                           entry order (see StopOffset)
' @param StopPrice      the trigger price for StopLossOrderTypeStop and StopLossOrderTypeStopLimit
' @param StopOffset     number of ticks to modify the StopPrice. For example, for a stop loss
'                       order 4 ticks below the entry order fill price, use StopLossOrderTypeAuto
'                       and set StopOffset to -4
' @param StopLimitPrice the limit price for StopLossOrderTypeStopLimit
' @param TargetOrderType the type of target order to place:
'           TargetOrderTypeNone             no target order
'           TargetOrderTypeLimit            use a limit order
'           TargetOrderTypeLimitIfTouched   use a limit if touched order
'           TargetOrderTypeMarketIfTouched  use a market if touched order
'           TargetOrderTypeBid              use a limit order at the current Bid (see TargetOffset)
'           TargetOrderTypeAsk              use a limit order at the current Ask (see TargetOffset)
'           TargetOrderTypeLast             use a limit order at the last Trade price (see TargetOffset)
'           TargetOrderTypeAuto             use a limit order at the fill price for the
'                                           entry order (see TargetOffset)
' @param TargetPrice    the limit price for the target order
' @param TargetOffset   number of ticks to modify the TargetPrice. For example, for a target
'                       order 10 ticks above the entry order fill price, use TargetOrderTypeAuto
'                       and set TargetOffset to 10
' @param TargetTriggerPrice  the trigger price for TargetOrderTypeLimitIfTouched and TargetOrderTypeMarketIfTouched
' @param CancelPrice    if the entry order is a limit order and market price rises to this
'                       price before the entry order is filled, then the order is cancelled.
'                       If the entry order is a stop or stop limit order and market price
'                       falls to this price before the entry order is filled, then the order
'                       is cancelled
' @param EntryTIF       specifies the time in force for the entry order
' @param StopTIF        specifies the time in force for the stop loss order
' @param TargetTIF      specifies the time in force for the target order
'
' @see
'
'@/
Public Function Buy(ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal buyPrice As Double, _
                    ByVal buyOffset As Long, _
                    ByVal TriggerPrice As Double, _
                    ByVal StopLossOrderType As StopLossOrderTypes, _
                    ByVal stopPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Long, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    Optional ByVal cancelPrice As Double, _
                    Optional ByVal entryTIF As OrderTIFs = OrderTIFDay, _
                    Optional ByVal stopTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
                    Optional ByVal targetTIF As OrderTIFs = OrderTIFGoodTillCancelled) As BracketOrder

Const ProcName As String = "Buy"

On Error GoTo Err

GLogLogger.Log pMsg:="Buy", _
    pMsgQualifier:=mTicker.Contract.Specifier.LocalSymbol & _
                " number=" & number & _
                "; EntryOrderType=" & GEntryOrderTypeToShortString(EntryOrderType) & _
                "; buyPrice=" & buyPrice & _
                "; buyOffset=" & buyOffset & _
                "; triggerPrice=" & TriggerPrice & _
                "; StopLossOrderType=" & GStopLossOrderTypeToShortString(StopLossOrderType) & _
                "; stopPrice=" & stopPrice & _
                "; stopOffset=" & stopOffset & _
                "; stopLimitPrice=" & stopLimitPrice & _
                "; TargetOrderType=" & GTargetOrderTypeToShortString(TargetOrderType) & _
                "; targetPrice=" & targetPrice & _
                "; targetOffset=" & targetOffset & _
                "; targetTriggerPrice=" & targetTriggerPrice & _
                "; cancelPrice=" & cancelPrice & _
                "; entryTIF=" & GOrderTIFToString(entryTIF) & _
                "; stopTIF=" & GOrderTIFToString(stopTIF) & _
                "; targetTIF=" & GOrderTIFToString(targetTIF), _
    pProcName:=ProcName, pModName:=ModuleName
                
Set Buy = CreateBracketOrder(OrderActions.ActionBuy, _
                                number, _
                                EntryOrderType, _
                                buyPrice, _
                                buyOffset, _
                                TriggerPrice, _
                                StopLossOrderType, _
                                stopPrice, _
                                stopOffset, _
                                stopLimitPrice, _
                                TargetOrderType, _
                                targetPrice, _
                                targetOffset, _
                                targetTriggerPrice)
                                
Buy.EntryOrder.TimeInForce = entryTIF
If Not Buy.Stoporder Is Nothing Then Buy.Stoporder.TimeInForce = stopTIF
If Not Buy.TargetOrder Is Nothing Then Buy.TargetOrder.TimeInForce = targetTIF

ExecuteBracketOrder Buy, True, cancelPrice

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Cancel() As Boolean
Dim lBracketOrder As BracketOrder

Const ProcName As String = "Cancel"

On Error GoTo Err

For Each lBracketOrder In mActiveBracketOrders
    If lBracketOrder.Cancel(evenIfFilled:=True) Then Cancel = True
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function CancelPending() As Boolean
Dim lBracketOrder As BracketOrder

Const ProcName As String = "CancelPending"

On Error GoTo Err

For Each lBracketOrder In mActiveBracketOrders
    If lBracketOrder.Cancel(evenIfFilled:=False) Then CancelPending = True
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub CloseOut(ByVal method As String)
Dim lBracketOrder As BracketOrder

Const ProcName As String = "CloseOut"

On Error GoTo Err

For Each lBracketOrder In mActiveBracketOrders
    lBracketOrder.CloseOut method
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Public Function CreateBuyBracketOrder( _
                ByRef number As Long, _
                ByVal EntryOrderType As EntryOrderTypes, _
                ByVal buyPrice As Double, _
                ByVal buyOffset As Long, _
                ByVal TriggerPrice As Double, _
                ByVal StopLossOrderType As StopLossOrderTypes, _
                ByVal stopPrice As Double, _
                ByVal stopOffset As Long, _
                ByVal stopLimitPrice As Long, _
                ByVal TargetOrderType As TargetOrderTypes, _
                ByVal targetPrice As Double, _
                ByVal targetOffset As Long, _
                ByVal targetTriggerPrice As Double) As BracketOrder
Const ProcName As String = "CreateBuyBracketOrder"

On Error GoTo Err

Set CreateBuyBracketOrder = CreateBracketOrder(OrderActions.ActionBuy, _
                                        number, _
                                        EntryOrderType, _
                                        buyPrice, _
                                        buyOffset, _
                                        TriggerPrice, _
                                        StopLossOrderType, _
                                        stopPrice, _
                                        stopOffset, _
                                        stopLimitPrice, _
                                        TargetOrderType, _
                                        targetPrice, _
                                        targetOffset, _
                                        targetTriggerPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function CreateBracketOrder(ByVal action As OrderActions, _
                    ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal price As Double, _
                    ByVal Offset As Long, _
                    ByVal TriggerPrice As Double, _
                    ByVal StopLossOrderType As StopLossOrderTypes, _
                    ByVal stopTriggerPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Double, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double) As BracketOrder
                    
Dim lBracketOrder As BracketOrder
Dim EntryOrder As Order
Dim Stoporder As Order
Dim TargetOrder As Order
Dim lEntryPrice As Double
Dim lTriggerPrice As Double
Dim lStopPrice As Double
Dim lTargetPrice As Double

Const ProcName As String = "CreateBracketOrder"

On Error GoTo Err

If Not IsReady Then Err.Raise ErrorCodes.ErrIllegalStateException, , "The service provider is not Ready"

If number <= 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The quantity must be greater than zero"

Select Case EntryOrderType
Case EntryOrderTypeMarket
Case EntryOrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
Case EntryOrderTypeMarketToLimit
Case EntryOrderTypeBid
Case EntryOrderTypeAsk
Case EntryOrderTypeLast
Case EntryOrderTypeLimit
Case EntryOrderTypeLimitOnOpen
Case EntryOrderTypeLimitOnClose
Case EntryOrderTypeLimitIfTouched
Case EntryOrderTypeStop
Case EntryOrderTypeStopLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid entry type"
End Select

Select Case StopLossOrderType
Case StopLossOrderTypeNone
Case StopLossOrderTypeStop
Case StopLossOrderTypeStopLimit
Case StopLossOrderTypeBid
Case StopLossOrderTypeAsk
Case StopLossOrderTypeLast
Case StopLossOrderTypeAuto
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid stop type"
End Select

Select Case TargetOrderType
Case TargetOrderTypeNone
Case TargetOrderTypeLimit
Case TargetOrderTypeLimitIfTouched
Case TargetOrderTypeMarketIfTouched
Case TargetOrderTypeBid
Case TargetOrderTypeAsk
Case TargetOrderTypeLast
Case TargetOrderTypeAuto
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid target type"
End Select

Set lBracketOrder = New BracketOrder
lBracketOrder.OrderContext = Me
lBracketOrder.Key = nextPlexKey
lBracketOrder.Ticker = mDataSource
lBracketOrder.OrderSubmitter = mOrderSubmitter
lBracketOrder.IsSimulated = mIsSimulated
lBracketOrder.PreventUnprotectedPosition = mPreventUnprotectedPositions

Set EntryOrder = lBracketOrder.CreateEntryOrder
EntryOrder.action = action
EntryOrder.Quantity = number

EntryOrder.OrderType = GEntryOrderTypeToOrderType(EntryOrderType)
Select Case EntryOrderType
Case EntryOrderTypeMarket
Case EntryOrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
    If TriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The trigger price must not be zero for this order type"
    
    EntryOrder.TriggerPrice = TriggerPrice
Case EntryOrderTypeMarketToLimit
Case EntryOrderTypeBid
    EntryOrder.LimitPrice = mDataSource.BidPrice + Offset * mDataSource.TickSize
Case EntryOrderTypeAsk
    EntryOrder.LimitPrice = mDataSource.AskPrice + Offset * mDataSource.TickSize
Case EntryOrderTypeLast
    EntryOrder.LimitPrice = mDataSource.TradePrice + Offset * mDataSource.TickSize
Case EntryOrderTypeLimit
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The entry price must not be zero for this order type"
    
    EntryOrder.LimitPrice = price + Offset * mDataSource.TickSize
Case EntryOrderTypeLimitOnOpen
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The entry price must not be zero for this order type"
    
    EntryOrder.LimitPrice = price + Offset * mDataSource.TickSize
Case EntryOrderTypeLimitOnClose
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The entry price must not be zero for this order type"
    
    EntryOrder.LimitPrice = price + Offset * mDataSource.TickSize
Case EntryOrderTypeLimitIfTouched
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The entry price must not be zero for this order type"
    If TriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The trigger price must not be zero for this order type"

    EntryOrder.LimitPrice = price + Offset * mDataSource.TickSize
    EntryOrder.TriggerPrice = TriggerPrice
Case EntryOrderTypeStop
    If TriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The trigger price must not be zero for this order type"
    
    EntryOrder.TriggerPrice = TriggerPrice
Case EntryOrderTypeStopLimit
    If TriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The trigger price must not be zero for this order type"
    
    EntryOrder.LimitPrice = price + Offset * mDataSource.TickSize
    EntryOrder.TriggerPrice = TriggerPrice
End Select

EntryOrder.Offset = Offset
lEntryPrice = EntryOrder.LimitPrice
lTriggerPrice = EntryOrder.TriggerPrice

If StopLossOrderType <> StopLossOrderTypeNone Then
    Set Stoporder = lBracketOrder.CreateStopOrder
    Stoporder.action = IIf(action = ActionBuy, OrderActions.ActionSell, OrderActions.ActionBuy)
    Stoporder.Quantity = number
    
    Stoporder.OrderType = GStopLossOrderTypeToOrderType(StopLossOrderType)
    Select Case StopLossOrderType
    Case StopLossOrderTypeStop
        If stopTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The stopTrigger price must not be zero for this order type"
        
        Stoporder.TriggerPrice = stopTriggerPrice + stopOffset * mDataSource.TickSize
    Case StopLossOrderTypeStopLimit
        If stopTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The stopTrigger price must not be zero for this order type"
        If stopLimitPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The stopLimit price must not be zero for this order type"
        
        Stoporder.TriggerPrice = stopTriggerPrice + stopOffset * mDataSource.TickSize
        Stoporder.LimitPrice = stopLimitPrice
    Case StopLossOrderTypeBid
        Stoporder.TriggerPrice = mDataSource.BidPrice + stopOffset * mDataSource.TickSize
        
    Case StopLossOrderTypeAsk
        Stoporder.TriggerPrice = mDataSource.AskPrice + stopOffset * mDataSource.TickSize
    Case StopLossOrderTypeLast
        Stoporder.TriggerPrice = mDataSource.TradePrice + stopOffset * mDataSource.TickSize
    Case StopLossOrderTypeAuto
        Stoporder.Offset = stopOffset * mDataSource.TickSize
    End Select
    lStopPrice = Stoporder.TriggerPrice
End If

If TargetOrderType <> TargetOrderTypeNone Then
    Set TargetOrder = lBracketOrder.CreateTargetOrder
    TargetOrder.action = IIf(action = ActionBuy, OrderActions.ActionSell, OrderActions.ActionBuy)
    TargetOrder.Quantity = number
    
    TargetOrder.OrderType = GTargetOrderTypeToOrderType(TargetOrderType)
    Select Case TargetOrderType
    Case TargetOrderTypeLimit
        If targetPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The target price must not be zero for this order type"
        
        TargetOrder.LimitPrice = targetPrice + targetOffset * mDataSource.TickSize
    Case TargetOrderTypeLimitIfTouched
        If targetPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The target price must not be zero for this order type"
        If targetTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The target trigger price must not be zero for this order type"
    
        TargetOrder.LimitPrice = targetPrice + Offset * mDataSource.TickSize
        TargetOrder.TriggerPrice = targetTriggerPrice
    Case TargetOrderTypeMarketIfTouched
        If targetTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The target trigger price must not be zero for this order type"
        
        TargetOrder.TriggerPrice = targetTriggerPrice
    Case TargetOrderTypeBid
        TargetOrder.LimitPrice = mDataSource.BidPrice + targetOffset * mDataSource.TickSize
    Case TargetOrderTypeAsk
        TargetOrder.LimitPrice = mDataSource.AskPrice + targetOffset * mDataSource.TickSize
    Case TargetOrderTypeLast
        TargetOrder.LimitPrice = mDataSource.TradePrice + targetOffset * mDataSource.TickSize
    Case TargetOrderTypeAuto
        TargetOrder.Offset = targetOffset * mDataSource.TickSize
    End Select

    lTargetPrice = TargetOrder.LimitPrice
End If

lBracketOrder.Checkpoint

Set CreateBracketOrder = lBracketOrder

GLogOrder "Create bracket order: " & _
            GOrderActionToString(action) & " " & number & " at " & _
            GOrderTypeToString(EntryOrder.OrderType) & " " & lEntryPrice & _
            IIf(lTriggerPrice <> 0, "(" & lTriggerPrice & ")", "") & _
            ", stop " & IIf(StopLossOrderType = StopLossOrderTypeAuto, "AUTO(" & stopOffset * mDataSource.TickSize & ")", lStopPrice) & _
            ", target " & IIf(TargetOrderType = TargetOrderTypeAuto, "AUTO(" & targetOffset * mDataSource.TickSize & ")", lTargetPrice) & _
            "(Curr price=" & mDataSource.FormatPrice(mDataSource.TradePrice) & _
            "; id=" & lBracketOrder.Key & ")", _
            mIsSimulated, _
            Me

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function CreateSellBracketOrder( _
                ByRef number As Long, _
                ByVal EntryOrderType As EntryOrderTypes, _
                ByVal buyPrice As Double, _
                ByVal buyOffset As Long, _
                ByVal TriggerPrice As Double, _
                ByVal StopLossOrderType As StopLossOrderTypes, _
                ByVal stopPrice As Double, _
                ByVal stopOffset As Long, _
                ByVal stopLimitPrice As Long, _
                ByVal TargetOrderType As TargetOrderTypes, _
                ByVal targetPrice As Double, _
                ByVal targetOffset As Long, _
                ByVal targetTriggerPrice As Double) As BracketOrder
Const ProcName As String = "CreateSellBracketOrder"

On Error GoTo Err

Set CreateSellBracketOrder = CreateBracketOrder(OrderActions.ActionSell, _
                                    number, _
                                    EntryOrderType, _
                                    buyPrice, _
                                    buyOffset, _
                                    TriggerPrice, _
                                    StopLossOrderType, _
                                    stopPrice, _
                                    stopOffset, _
                                    stopLimitPrice, _
                                    TargetOrderType, _
                                    targetPrice, _
                                    targetOffset, _
                                    targetTriggerPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub ExecuteBracketOrder(ByVal pBracketOrder As BracketOrder, _
                    Optional checkMoneyManagement As Boolean, _
                    Optional cancelPrice As Double)
                    
Dim msg As String
Dim origQuantity As Long
Dim NewQuantity As Long

Const ProcName As String = "ExecuteBracketOrder"

On Error GoTo Err

If pBracketOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "No bracket order supplied"

If checkMoneyManagement And Not mIsSimulated Then
    origQuantity = pBracketOrder.EntryOrder.Quantity
    NewQuantity = origQuantity
    If Not mOrderAuthoriser.AllowOrder( _
                            IIf(pBracketOrder.EntryOrder.action = ActionBuy, True, False), _
                            NewQuantity, _
                            pBracketOrder.Risk, _
                            msg) _
    Then
        GLogOrder "Rejected by money management: " & msg & " " & _
                            "(Curr price=" & mDataSource.FormatPrice(mDataSource.TradePrice) & _
                            "; id=" & pBracketOrder.Key & ")", _
                            mIsSimulated, _
                            Me
        Set pBracketOrder = Nothing
        Exit Sub
    End If
    If NewQuantity <> origQuantity Then
        pBracketOrder.NewQuantity = NewQuantity
        pBracketOrder.Update
    End If
End If

GLogOrder "Execute bracket order " & _
            "(Curr price=" & mDataSource.FormatPrice(mDataSource.TradePrice) & _
            "; id=" & pBracketOrder.Key & ")", _
            mIsSimulated, _
            Me

mActiveBracketOrders.Add pBracketOrder
mAccumulatedBracketOrders.Add pBracketOrder
pBracketOrder.Execute cancelPrice

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function Finish()
Dim op As BracketOrder

Const ProcName As String = "Finish"

On Error GoTo Err

Set mDataSource = Nothing
For Each op In mActiveBracketOrders
    op.Finish
Next
Set mActiveBracketOrders = Nothing

For Each op In mClosedBracketOrders
    op.Finish
Next
Set mClosedBracketOrders = Nothing

Set mOrderSubmitter = Nothing

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function GetAccountBalance() As Double
Const ProcName As String = "GetAccountBalance"

On Error GoTo Err

GetAccountBalance = 10000                       '!!!!!!!!!!!!!!!!!!!!!!!!! need to fix this up

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Initialise( _
                ByVal pName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pOrderAuthoriser As IOrderAuthoriser, _
                ByVal pAccumulatedBracketOrders As Orders, _
                ByVal pAccumulatedOrders As Orders)
mName = pName
mIsSimulated = pIsSimulated
Set mDataSource = pDataSource
Set mOrderSubmitter = pOrderSubmitter
Set pOrderAuthoriser = pOrderAuthoriser
Set mAccumulatedBracketOrders = pAccumulatedBracketOrders
Set mAccumulatedOrders = pAccumulatedOrders
End Sub

Public Function IsAttributeSupported( _
                ByVal attributeId As OrderAttributes) As Boolean
IsAttributeSupported = (mPermittedOrderAttributes And attributeId)
End Function

Public Function IsOrderTifValueSupported( _
                ByVal value As OrderTIFs) As Boolean
IsOrderTifValueSupported = (mPermittedOrderTifs And value)
End Function

Public Function IsOrderTypeValueSupported( _
                ByVal value As OrderTypes) As Boolean
If value = OrderTypeAutoLimit Or _
    value = OrderTypeAutoStop _
Then
    IsOrderTypeValueSupported = True
Else
    IsOrderTypeValueSupported = (mPermittedOrderTypes And value)
End If
End Function

Public Function IsStopTriggerMethodValueSupported( _
                ByVal value As StopTriggerMethods) As Boolean
IsStopTriggerMethodValueSupported = (mPermittedStopTriggerMethods And value)
End Function

Friend Function NewOrder() As Order
Dim lOrder As Order
Const ProcName As String = "NewOrder"

On Error GoTo Err

Set lOrder = New Order
With lOrder
    .Id = GenerateGUIDString
    .IsSimulated = mIsSimulated
    .OrderContext = Me
    .Ticker = mDataSource
    If IsAttributeSupported(OrderAttributes.OrderAttAllOrNone) Then .AllOrNone = mAllOrNone
    If IsAttributeSupported(OrderAttributes.OrderAttBlockOrder) Then .BlockOrder = mBlockOrder
    If IsAttributeSupported(OrderAttributes.OrderAttDiscretionaryAmount) Then .DiscretionaryAmount = mDiscretionaryAmount
    If IsAttributeSupported(OrderAttributes.OrderAttDisplaySize) Then .DisplaySize = mDisplaySize
    If IsAttributeSupported(OrderAttributes.OrderAttETradeOnly) Then .ETradeOnly = mETradeOnly
    If IsAttributeSupported(OrderAttributes.OrderAttFirmQuoteOnly) Then .FirmQuoteOnly = mFirmQuoteOnly
    If IsAttributeSupported(OrderAttributes.OrderAttGoodAfterTime) Then .GoodAfterTime = mGoodAfterTime
    If IsAttributeSupported(OrderAttributes.OrderAttGoodAfterTimeTZ) Then .GoodAfterTimeTZ = mGoodAfterTimeTZ
    If IsAttributeSupported(OrderAttributes.OrderAttGoodTillDate) Then .GoodTillDate = mGoodTillDate
    If IsAttributeSupported(OrderAttributes.OrderAttGoodTillDateTZ) Then .GoodTillDateTZ = mGoodTillDateTZ
    If IsAttributeSupported(OrderAttributes.OrderAttHidden) Then .Hidden = mHidden
    If IsAttributeSupported(OrderAttributes.OrderAttIgnoreRTH) Then .IgnoreRegularTradingHours = mIgnoreRegularTradingHours
    If IsAttributeSupported(OrderAttributes.OrderAttLimitPrice) Then .LimitPrice = mLimitPrice
    If IsAttributeSupported(OrderAttributes.OrderAttMinimumQuantity) Then .MinimumQuantity = mMinimumQuantity
    If IsAttributeSupported(OrderAttributes.OrderAttNBBOPriceCap) Then .NbboPriceCap = mNBBOPriceCap
    If mOrderType <> OrderTypes.OrderTypeNone Then
        If IsAttributeSupported(OrderAttributes.OrderAttOrderType) Then .OrderType = mOrderType
    End If
    If IsAttributeSupported(OrderAttributes.OrderAttOrigin) Then .Origin = mOrigin
    If IsAttributeSupported(OrderAttributes.OrderAttOverrideConstraints) Then .OverrideConstraints = mOverrideConstraints
    If IsAttributeSupported(OrderAttributes.OrderAttPercentOffset) Then .PercentOffset = mPercentOffset
    If IsAttributeSupported(OrderAttributes.OrderAttQuantity) Then .Quantity = mQuantity
    If IsAttributeSupported(OrderAttributes.OrderAttSettlingFirm) Then .SettlingFirm = mSettlingFirm
    
    If mStopTriggerMethod <> StopTriggerMethods.StopTriggerNone Then
        If IsAttributeSupported(OrderAttributes.OrderAttStopTriggerMethod) Then .StopTriggerMethod = mStopTriggerMethod
    End If
    If IsAttributeSupported(OrderAttributes.OrderAttSweepToFill) Then .SweepToFill = mSweepToFill
    If mTimeInForce <> OrderTIFs.TIFNone Then
        If IsAttributeSupported(OrderAttributes.OrderAttTimeInForce) Then .TimeInForce = mTimeInForce
    End If
    If IsAttributeSupported(OrderAttributes.OrderAttTriggerPrice) Then .TriggerPrice = mTriggerPrice
End With

mAccumulatedOrders.Add lOrder

Set NewOrder = lOrder

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub NotifyEndOfTradingSession( _
                ByVal pTimestamp As Date)
Const ProcName As String = "NotifyEndOfTradingSession"

On Error GoTo Err

Dim op As BracketOrder

For Each op In mActiveBracketOrders
    op.NotifyEndOfTradingSession pTimestamp
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub NotifyPositionCloseout( _
                ByVal pPrice As Double, _
                ByVal pTimestamp As Date)
Const ProcName As String = "NotifyCloseout"

On Error GoTo Err

Dim op As BracketOrder

For Each op In mClosedBracketOrders
    op.NotifyCloseout pPrice, pTimestamp
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function NotifyBracketOrderClosed( _
                ByVal pBracketOrder As BracketOrder)
Const ProcName As String = "NotifyBracketOrderClosed"

On Error GoTo Err

On Error Resume Next
mActiveBracketOrders.Remove pBracketOrder
On Error GoTo Err

If Not mClosedBracketOrders.Contains(pBracketOrder.Key) Then
    mClosedBracketOrders.Add pBracketOrder
    mCompletedSize = mCompletedSize + pBracketOrder.size
    ' now check if there are any active bracket orders left -
    ' if not, need to tell the Position Manager
    If mActiveBracketOrders.Count = 0 Then
        mPositionManager.NotifyOrderContextInactive Me
    End If
End If
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Sell(ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal sellPrice As Double, _
                    ByVal sellOffset As Long, _
                    ByVal TriggerPrice As Double, _
                    ByVal StopLossOrderType As StopLossOrderTypes, _
                    ByVal stopPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Long, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    Optional ByVal cancelPrice As Double, _
                    Optional ByVal entryTIF As OrderTIFs = OrderTIFDay, _
                    Optional ByVal stopTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
                    Optional ByVal targetTIF As OrderTIFs = OrderTIFGoodTillCancelled) As BracketOrder

Const ProcName As String = "Sell"

On Error GoTo Err

GLogLogger.Log pMsg:="Sell", _
    pMsgQualifier:=mTicker.Contract.Specifier.LocalSymbol & _
                " number=" & number & _
                "; EntryOrderType=" & GEntryOrderTypeToShortString(EntryOrderType) & _
                "; sellPrice=" & sellPrice & _
                "; sellOffset=" & sellOffset & _
                "; triggerPrice=" & TriggerPrice & _
                "; StopLossOrderType=" & GStopLossOrderTypeToShortString(StopLossOrderType) & _
                "; stopPrice=" & stopPrice & _
                "; stopOffset=" & stopOffset & _
                "; stopLimitPrice=" & stopLimitPrice & _
                "; TargetOrderType=" & GTargetOrderTypeToShortString(TargetOrderType) & _
                "; targetPrice=" & targetPrice & _
                "; targetOffset=" & targetOffset & _
                "; targetTriggerPrice=" & targetTriggerPrice & _
                "; cancelPrice=" & cancelPrice & _
                "; entryTIF=" & GOrderTIFToString(entryTIF) & _
                "; stopTIF=" & GOrderTIFToString(stopTIF) & _
                "; targetTIF=" & GOrderTIFToString(targetTIF) & _
                "; stopStrategy=" & TypeName(StopStrategy) & _
                "; exitStrategy=" & TypeName(ExitStrategy) & _
                "; scalingStrategy=" & TypeName(ScalingStrategy), _
    pProcName:=ProcName, pModName:=ModuleName
                
Set Sell = CreateBracketOrder(OrderActions.ActionSell, _
                                number, _
                                EntryOrderType, _
                                sellPrice, _
                                sellOffset, _
                                TriggerPrice, _
                                StopLossOrderType, _
                                stopPrice, _
                                stopOffset, _
                                stopLimitPrice, _
                                TargetOrderType, _
                                targetPrice, _
                                targetOffset, _
                                targetTriggerPrice)

Sell.EntryOrder.TimeInForce = entryTIF
If Not Sell.Stoporder Is Nothing Then Sell.Stoporder.TimeInForce = stopTIF
If Not Sell.TargetOrder Is Nothing Then Sell.TargetOrder.TimeInForce = targetTIF

ExecuteBracketOrder Sell, True, cancelPrice, StopStrategy, ExitStrategy, ScalingStrategy

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attributeValueNotSupported( _
                ByVal pAttribute As OrderAttributes, _
                ByVal value As Variant)
Const ProcName As String = "attributeValueNotSupported"

On Error GoTo Err

Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value " & CStr(value) & " for attribute " & GOrderAttributeToString(pAttribute) & " is not supported by the service provider"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub checkAttributeSupported( _
                ByVal pAttribute As OrderAttributes)
Const ProcName As String = "checkAttributeSupported"

On Error GoTo Err

If Not IsAttributeSupported(pAttribute) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "The " & GOrderAttributeToString(pAttribute) & " attribute is not supported by the service provider"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub checkReady()
Const ProcName As String = "checkReady"

On Error GoTo Err

If IsReady Then
    mPermittedOrderAttributes = mOrderSubmitter.PermittedOrderAttributes
    mPermittedOrderTifs = mOrderSubmitter.PermittedOrderTifs
    mPermittedOrderTypes = mOrderSubmitter.PermittedOrderTypes
    mPermittedStopTriggerMethods = mOrderSubmitter.PermittedStopTriggerMethods
    RaiseEvent Ready
Else
    RaiseEvent NotReady
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function nextPlexKey() As String
Static currID As Long
Const ProcName As String = "nextPlexKey"

On Error GoTo Err

currID = currID + 1
nextPlexKey = gTB.SessionId & "/" & currID

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

