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

Event ready()
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

Private mName As String

Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1
Private mPositionManager As PositionManager

Private mTickSize As Double

Private mIsSimulated As Boolean

Private mActiveOrderPlexes As OrderPlexes
    ' this contains all order plexes for which one or more orders are
    ' not yet complete (or not yet cancelled)
    
Private mClosedOrderPlexes As OrderPlexes
    ' this contains all order plexes for which all orders have completed (or
    ' been cancelled)
    
Private mCompletedSize As Long
    ' the total Size for all completed order plexes (may be non-zero if one
    ' or more completed order plexes has no target or stop orders)
    
Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter
Private mPermittedOrderAttributes As Long
Private mPermittedOrderTifs As Long
Private mPermittedOrderTypes As Long
Private mPermittedStopTriggerMethods As Long

Private mEnabled As Boolean
Private mIsReady As Boolean

Private mOrigin As OrderOrigins
Private mBlockOrder As Boolean
Private mSweepToFill As Boolean
Private mDisplaySize As Long
Private mIgnoreRegularTradingHours As Boolean
Private mHidden As Boolean
Private mDiscretionaryAmount As Double
Private mGoodAfterTime As Date
Private mGoodAfterTimeTZ    As String
Private mGoodTillDate   As Date
Private mGoodTillDateTZ As String
'Private mRule80A               ' not implemented
Private mSettlingFirm As String
Private mAllOrNone As Boolean
Private mMinimumQuantity As Long
Private mPercentOffset As Double
Private mETradeOnly As Boolean
Private mFirmQuoteOnly  As Boolean
Private mNBBOPriceCap As Double
Private mOverrideConstraints  As Boolean
Private mLimitPrice As Double
Private mOrderType As OrderTypes
Private mQuantity As Long
Private mTimeInForce As OrderTifs
Private mTriggerPrice As Double
Private mStopTriggerMethod As StopTriggerMethods

Private mPreventUnprotectedPositions As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mActiveOrderPlexes = New OrderPlexes
Set mClosedOrderPlexes = New OrderPlexes
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderContext terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_StateChange(ev As TWUtilities30.StateChangeEvent)
Const ProcName As String = "mTicker_StateChange"
Dim failpoint As String
On Error GoTo Err

If ev.State = TickerStates.TickerStateReady Then
    mTickSize = mTicker.Contract.TickSize
End If

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ActiveOrderPlexes() As OrderPlexes
Set ActiveOrderPlexes = mActiveOrderPlexes
End Property

Public Property Get ActiveSize() As Long
Dim lOrderPlex As orderPlex

Const ProcName As String = "ActiveSize"
Dim failpoint As Long
On Error GoTo Err

For Each lOrderPlex In mActiveOrderPlexes
    ActiveSize = ActiveSize + lOrderPlex.size
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let AllOrNone(ByVal value As Boolean)
Const ProcName As String = "AllOrNone"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttAllOrNone
mAllOrNone = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mAllOrNone
End Property

Public Property Let BlockOrder(ByVal value As Boolean)
Const ProcName As String = "BlockOrder"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttBlockOrder
mBlockOrder = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mBlockOrder
End Property

Public Property Get ClosedOrderPlexes() As OrderPlexes
Set ClosedOrderPlexes = mClosedOrderPlexes
End Property

Public Property Get Complete() As Boolean
Dim op As orderPlex
Const ProcName As String = "Complete"
Dim failpoint As Long
On Error GoTo Err

For Each op In mActiveOrderPlexes
    If Not op.IsSimulated Then Complete = False: Exit Property
Next
Complete = True

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get CompletedSize() As Long
CompletedSize = mCompletedSize
End Property

Public Property Let DisplaySize(ByVal value As Long)
Const ProcName As String = "DisplaySize"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttDisplaySize
mDisplaySize = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Friend Property Let Enabled(ByVal value As Boolean)
Const ProcName As String = "Enabled"
Dim failpoint As Long
On Error GoTo Err

mEnabled = value
checkReady

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let ETradeOnly(ByVal value As Boolean)
Const ProcName As String = "ETradeOnly"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttETradeOnly
mETradeOnly = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get ETradeOnly() As Boolean
ETradeOnly = mETradeOnly
End Property

Public Property Let FirmQuoteOnly(ByVal value As Boolean)
Const ProcName As String = "FirmQuoteOnly"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttFirmQuoteOnly
mFirmQuoteOnly = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let GoodAfterTime(ByVal value As Date)
Const ProcName As String = "GoodAfterTime"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttGoodAfterTime
mGoodAfterTime = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal value As String)
Const ProcName As String = "GoodAfterTimeTZ"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttGoodAfterTimeTZ
mGoodAfterTimeTZ = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal value As Date)
Const ProcName As String = "GoodTillDate"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttGoodTillDate
mGoodTillDate = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mGoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal value As String)
Const ProcName As String = "GoodTillDateTZ"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttGoodTillDateTZ
mGoodTillDateTZ = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let Hidden(ByVal value As Boolean)
Const ProcName As String = "Hidden"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttHidden
mHidden = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Public Property Let IgnoreRegularTradingHours(ByVal value As Boolean)
Const ProcName As String = "IgnoreRegularTradingHours"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttIgnoreRTH
mIgnoreRegularTradingHours = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Friend Property Let IsReady(ByVal value As Boolean)
Const ProcName As String = "IsReady"
Dim failpoint As Long
On Error GoTo Err

mIsReady = value
checkReady

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get IsReady() As Boolean
IsReady = (mIsReady And mEnabled)
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
mIsSimulated = value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Let LimitPrice(ByVal value As Double)
Const ProcName As String = "LimitPrice"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttLimitPrice
mLimitPrice = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Public Property Let MinimumQuantity(ByVal value As Long)
Const ProcName As String = "MinimumQuantity"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttMinimumQuantity
mMinimumQuantity = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mMinimumQuantity
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Friend Property Get name() As String
name = mName
End Property

Public Property Let NbboPriceCap(ByVal value As Double)
Const ProcName As String = "NbboPriceCap"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttNBBOPriceCap
mNBBOPriceCap = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Public Property Let OrderType(ByVal value As OrderTypes)
Const ProcName As String = "OrderType"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttOrderType
If Not IsOrderTypeValueSupported(value) Then attributeValueNotSupported OrderAttributeIds.OrderAttOrderType, value
mOrderType = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let Origin(ByVal value As OrderOrigins)
Const ProcName As String = "Origin"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttOrigin
mOrigin = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Origin() As OrderOrigins
Origin = mOrigin
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Set mOrderSubmitter = value
End Property

Public Property Let OverrideConstraints(ByVal value As Boolean)
Const ProcName As String = "OverrideConstraints"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttOverrideConstraints
mOverrideConstraints = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mOverrideConstraints
End Property

Public Property Get PendingSize() As Long
Dim lOrderPlex As orderPlex
Const ProcName As String = "PendingSize"
Dim failpoint As Long
On Error GoTo Err

For Each lOrderPlex In mActiveOrderPlexes
    PendingSize = PendingSize + lOrderPlex.PendingSize
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let PercentOffset(ByVal value As Double)
Const ProcName As String = "PercentOffset"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttPercentOffset
mPercentOffset = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mPercentOffset
End Property

Public Property Get PermittedOrderTifs() As Long
PermittedOrderTifs = mPermittedOrderTifs
End Property

Public Property Get permittedOrderTypes() As Long
permittedOrderTypes = mPermittedOrderTypes
End Property

Public Property Get permittedStopTriggerMethods() As Long
permittedStopTriggerMethods = mPermittedStopTriggerMethods
End Property

Friend Property Let PositionManager(ByVal value As PositionManager)
Set mPositionManager = value
End Property

Public Property Let PreventUnprotectedPositions( _
                ByVal value As Boolean)
mPreventUnprotectedPositions = value
End Property

Public Property Get PreventUnprotectedPositions() As Boolean
PreventUnprotectedPositions = mPreventUnprotectedPositions
End Property

Public Property Let Quantity(ByVal value As Long)
Const ProcName As String = "Quantity"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttQuantity
mQuantity = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Public Property Get Risk() As Currency
Dim lOrderPlex As orderPlex
Const ProcName As String = "Risk"
Dim failpoint As Long
On Error GoTo Err

For Each lOrderPlex In mActiveOrderPlexes
    If lOrderPlex.Risk = MaxCurrency Then
        Risk = MaxCurrency
        Exit For
    End If
    Risk = Risk + lOrderPlex.Risk
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let SettlingFirm(ByVal value As String)
Const ProcName As String = "SettlingFirm"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttSettlingFirm
mSettlingFirm = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSettlingFirm
End Property

Public Property Get size() As Long
Dim lOrderPlex As orderPlex

Const ProcName As String = "Size"
Dim failpoint As Long
On Error GoTo Err

size = mCompletedSize

For Each lOrderPlex In mActiveOrderPlexes
    size = size + lOrderPlex.size
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let StopTriggerMethod(ByVal value As StopTriggerMethods)
Const ProcName As String = "StopTriggerMethod"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttStopTriggerMethod
If Not IsStopTriggerMethodValueSupported(value) Then attributeValueNotSupported OrderAttributeIds.OrderAttStopTriggerMethod, value
mStopTriggerMethod = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get StopTriggerMethod() As StopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(ByVal value As Boolean)
Const ProcName As String = "SweepToFill"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttSweepToFill
mSweepToFill = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Const ProcName As String = "Ticker"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

Public Property Let TimeInForce(ByVal value As OrderTifs)
Const ProcName As String = "TimeInForce"
Dim failpoint As Long
On Error GoTo Err

If Not GIsValidTIF(value) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "timeInForce", _
            "Value must be a member of the OrderTifs enum"
End If

checkAttributeSupported OrderAttributeIds.OrderAttTimeInForce
If Not IsOrderTifValueSupported(value) Then attributeValueNotSupported OrderAttributeIds.OrderAttTimeInForce, value
mTimeInForce = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get TimeInForce() As OrderTifs
TimeInForce = mTimeInForce
End Property

Public Property Let TriggerPrice(ByVal value As Double)
Const ProcName As String = "TriggerPrice"
Dim failpoint As Long
On Error GoTo Err

checkAttributeSupported OrderAttributeIds.OrderAttTriggerPrice
mTriggerPrice = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ActivateOrderPlex( _
                ByVal pOrderPlex As orderPlex)
Const ProcName As String = "ActivateOrderPlex"
Dim failpoint As String
On Error GoTo Err

pOrderPlex.Ticker = mTicker
pOrderPlex.orderSubmitter = mOrderSubmitter
mActiveOrderPlexes.Add pOrderPlex

pOrderPlex.Activate

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

''
' Creates a buy<code>OrderPlex</code>
'
' @return the <code>OrderPlex</code> object created
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
' @param StopOrderType  the type of stop loss order to place:
'           StopOrderTypeNone               no stop loss order
'           StopOrderTypeStop               use a stop order
'           StopOrderTypeStopLimit          use a stop limit order
'           StopOrderTypeBid                use a stop order at the current Bid (see StopOffset)
'           StopOrderTypeAsk                use a stop order at the current Ask (see StopOffset)
'           StopOrderTypeLast               use a stop order at the last Trade price (see StopOffset)
'           StopOrderTypeAuto               use a stop order at the fill price for the
'                                           entry order (see StopOffset)
' @param StopPrice      the trigger price for StopOrderTypeStop and StopOrderTypeStopLimit
' @param StopOffset     number of ticks to modify the StopPrice. For example, for a stop loss
'                       order 4 ticks below the entry order fill price, use StopOrderTypeAuto
'                       and set StopOffset to -4
' @param StopLimitPrice the limit price for StopOrderTypeStopLimit
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
' @param StopStrategy   an object that manages the stop loss order
' @param ExitStrategy   an object that manages the exit strategy
' @param ScalingStrategy an object that manages scaling into the position (ie placing further
'                       orders to increase the position)
'
' @see
'
'@/
Public Function Buy(ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal buyPrice As Double, _
                    ByVal buyOffset As Long, _
                    ByVal TriggerPrice As Double, _
                    ByVal StopOrderType As StopOrderTypes, _
                    ByVal stopPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Long, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    Optional ByVal cancelPrice As Double, _
                    Optional ByVal entryTIF As OrderTifs = TIFDay, _
                    Optional ByVal stopTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal targetTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal StopStrategy As StopManagementStrategy, _
                    Optional ByVal ExitStrategy As PositionManagementStrategy, _
                    Optional ByVal ScalingStrategy As PositionManagementStrategy) As orderPlex

Const ProcName As String = "Buy"
Dim failpoint As Long
On Error GoTo Err

GLog pMsg:="Buy", _
    pMsgQualifier:=mTicker.Contract.specifier.LocalSymbol & _
                " number=" & number & _
                "; EntryOrderType=" & GEntryOrderTypeToShortString(EntryOrderType) & _
                "; buyPrice=" & buyPrice & _
                "; buyOffset=" & buyOffset & _
                "; triggerPrice=" & TriggerPrice & _
                "; StopOrderType=" & GStopOrderTypeToShortString(StopOrderType) & _
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
    pProcName:="Buy", pProjName:=ProjectName, pModName:=ModuleName
                
Set Buy = CreateOrderPlex(OrderActions.ActionBuy, _
                                number, _
                                EntryOrderType, _
                                buyPrice, _
                                buyOffset, _
                                TriggerPrice, _
                                StopOrderType, _
                                stopPrice, _
                                stopOffset, _
                                stopLimitPrice, _
                                TargetOrderType, _
                                targetPrice, _
                                targetOffset, _
                                targetTriggerPrice)
                                
Buy.EntryOrder.TimeInForce = entryTIF
If Not Buy.stoporder Is Nothing Then Buy.stoporder.TimeInForce = stopTIF
If Not Buy.TargetOrder Is Nothing Then Buy.TargetOrder.TimeInForce = targetTIF

ExecuteOrderPlex Buy, True, cancelPrice, StopStrategy, ExitStrategy, ScalingStrategy

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function Cancel() As Boolean
Dim lOrderPlex As orderPlex

Const ProcName As String = "Cancel"
Dim failpoint As Long
On Error GoTo Err

For Each lOrderPlex In mActiveOrderPlexes
    If lOrderPlex.Cancel(evenIfFilled:=True) Then Cancel = True
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function CancelPending() As Boolean
Dim lOrderPlex As orderPlex

Const ProcName As String = "CancelPending"
Dim failpoint As Long
On Error GoTo Err

For Each lOrderPlex In mActiveOrderPlexes
    If lOrderPlex.Cancel(evenIfFilled:=False) Then CancelPending = True
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub CloseOut(ByVal method As String)
Dim lOrderPlex As orderPlex

Const ProcName As String = "CloseOut"
Dim failpoint As Long
On Error GoTo Err

For Each lOrderPlex In mActiveOrderPlexes
    lOrderPlex.CloseOut method
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Public Function CreateBuyOrderPlex( _
                ByRef number As Long, _
                ByVal EntryOrderType As EntryOrderTypes, _
                ByVal buyPrice As Double, _
                ByVal buyOffset As Long, _
                ByVal TriggerPrice As Double, _
                ByVal StopOrderType As StopOrderTypes, _
                ByVal stopPrice As Double, _
                ByVal stopOffset As Long, _
                ByVal stopLimitPrice As Long, _
                ByVal TargetOrderType As TargetOrderTypes, _
                ByVal targetPrice As Double, _
                ByVal targetOffset As Long, _
                ByVal targetTriggerPrice As Double) As orderPlex
Const ProcName As String = "CreateBuyOrderPlex"
Dim failpoint As Long
On Error GoTo Err

Set CreateBuyOrderPlex = CreateOrderPlex(OrderActions.ActionBuy, _
                                        number, _
                                        EntryOrderType, _
                                        buyPrice, _
                                        buyOffset, _
                                        TriggerPrice, _
                                        StopOrderType, _
                                        stopPrice, _
                                        stopOffset, _
                                        stopLimitPrice, _
                                        TargetOrderType, _
                                        targetPrice, _
                                        targetOffset, _
                                        targetTriggerPrice)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Function CreateOrderPlex(ByVal action As OrderActions, _
                    ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal price As Double, _
                    ByVal Offset As Long, _
                    ByVal TriggerPrice As Double, _
                    ByVal StopOrderType As StopOrderTypes, _
                    ByVal stopTriggerPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Double, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double) As orderPlex
                    
Dim lOrderPlex As orderPlex
Dim EntryOrder As Order
Dim stoporder As Order
Dim TargetOrder As Order
Dim lEntryPrice As Double
Dim lTriggerPrice As Double
Dim lStopPrice As Double
Dim lTargetPrice As Double

Const ProcName As String = "CreateOrderPlex"
Dim failpoint As Long
On Error GoTo Err

If Not IsReady Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "OrderContext::CreateOrderPlex", _
                            "The service provider is not Ready"

If number <= 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            "OrderContext::CreateOrderPlex", _
                            "The quantity must be greater than zero"

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
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "OrderContext::CreateOderPlex", _
    "Invalid entry type"
End Select

Select Case StopOrderType
Case StopOrderTypeNone
Case StopOrderTypeStop
Case StopOrderTypeStopLimit
Case StopOrderTypeBid
Case StopOrderTypeAsk
Case StopOrderTypeLast
Case StopOrderTypeAuto
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "OrderContext::CreateOderPlex", _
    "Invalid stop type"
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
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "OrderContext::CreateOderPlex", _
    "Invalid target type"
End Select

Set lOrderPlex = New orderPlex
lOrderPlex.OrderContext = Me
lOrderPlex.Key = nextPlexKey
lOrderPlex.PositionManager = mPositionManager
lOrderPlex.Ticker = mTicker
lOrderPlex.orderSubmitter = mOrderSubmitter
lOrderPlex.IsSimulated = mIsSimulated
lOrderPlex.PreventUnprotectedPosition = mPreventUnprotectedPositions

Set EntryOrder = lOrderPlex.CreateEntryOrder
EntryOrder.action = action
EntryOrder.Quantity = number

EntryOrder.OrderType = GEntryOrderTypeToOrderType(EntryOrderType)
Select Case EntryOrderType
Case EntryOrderTypeMarket
Case EntryOrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
    If TriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"
    
    EntryOrder.TriggerPrice = TriggerPrice
Case EntryOrderTypeMarketToLimit
Case EntryOrderTypeBid
    EntryOrder.LimitPrice = mTicker.BidPrice + Offset * mTickSize
Case EntryOrderTypeAsk
    EntryOrder.LimitPrice = mTicker.AskPrice + Offset * mTickSize
Case EntryOrderTypeLast
    EntryOrder.LimitPrice = mTicker.TradePrice + Offset * mTickSize
Case EntryOrderTypeLimit
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    EntryOrder.LimitPrice = price + Offset * mTickSize
Case EntryOrderTypeLimitOnOpen
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    EntryOrder.LimitPrice = price + Offset * mTickSize
Case EntryOrderTypeLimitOnClose
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    EntryOrder.LimitPrice = price + Offset * mTickSize
Case EntryOrderTypeLimitIfTouched
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    If TriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"

    EntryOrder.LimitPrice = price + Offset * mTickSize
    EntryOrder.TriggerPrice = TriggerPrice
Case EntryOrderTypeStop
    If TriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"
    
    EntryOrder.TriggerPrice = TriggerPrice
Case EntryOrderTypeStopLimit
    If TriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"
    
    EntryOrder.LimitPrice = price + Offset * mTickSize
    EntryOrder.TriggerPrice = TriggerPrice
End Select

EntryOrder.Offset = Offset
lEntryPrice = EntryOrder.LimitPrice
lTriggerPrice = EntryOrder.TriggerPrice

If StopOrderType <> StopOrderTypeNone Then
    Set stoporder = lOrderPlex.createStopOrder
    stoporder.action = IIf(action = ActionBuy, OrderActions.ActionSell, OrderActions.ActionBuy)
    stoporder.Quantity = number
    
    stoporder.OrderType = GStopOrderTypeToOrderType(StopOrderType)
    Select Case StopOrderType
    Case StopOrderTypeStop
        If stopTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The stopTrigger price must not be zero for this order type"
        
        stoporder.TriggerPrice = stopTriggerPrice + stopOffset * mTickSize
    Case StopOrderTypeStopLimit
        If stopTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The stopTrigger price must not be zero for this order type"
        If stopLimitPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The stopLimit price must not be zero for this order type"
        
        stoporder.TriggerPrice = stopTriggerPrice + stopOffset * mTickSize
        stoporder.LimitPrice = stopLimitPrice
    Case StopOrderTypeBid
        stoporder.TriggerPrice = mTicker.BidPrice + stopOffset * mTickSize
        
    Case StopOrderTypeAsk
        stoporder.TriggerPrice = mTicker.AskPrice + stopOffset * mTickSize
    Case StopOrderTypeLast
        stoporder.TriggerPrice = mTicker.TradePrice + stopOffset * mTickSize
    Case StopOrderTypeAuto
        stoporder.Offset = stopOffset * mTickSize
    End Select
    lStopPrice = stoporder.TriggerPrice
End If

If TargetOrderType <> TargetOrderTypeNone Then
    Set TargetOrder = lOrderPlex.createTargetOrder
    TargetOrder.action = IIf(action = ActionBuy, OrderActions.ActionSell, OrderActions.ActionBuy)
    TargetOrder.Quantity = number
    
    TargetOrder.OrderType = GTargetOrderTypeToOrderType(TargetOrderType)
    Select Case TargetOrderType
    Case TargetOrderTypeLimit
        If targetPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The target price must not be zero for this order type"
        
        TargetOrder.LimitPrice = targetPrice + targetOffset * mTickSize
    Case TargetOrderTypeLimitIfTouched
        If targetPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The target price must not be zero for this order type"
        If targetTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The target trigger price must not be zero for this order type"
    
        TargetOrder.LimitPrice = targetPrice + Offset * mTickSize
        TargetOrder.TriggerPrice = targetTriggerPrice
    Case TargetOrderTypeMarketIfTouched
        If targetTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The target trigger price must not be zero for this order type"
        
        TargetOrder.TriggerPrice = targetTriggerPrice
    Case TargetOrderTypeBid
        TargetOrder.LimitPrice = mTicker.BidPrice + targetOffset * mTickSize
    Case TargetOrderTypeAsk
        TargetOrder.LimitPrice = mTicker.AskPrice + targetOffset * mTickSize
    Case TargetOrderTypeLast
        TargetOrder.LimitPrice = mTicker.TradePrice + targetOffset * mTickSize
    Case TargetOrderTypeAuto
        TargetOrder.Offset = targetOffset * mTickSize
    End Select

    lTargetPrice = TargetOrder.LimitPrice
End If

lOrderPlex.Checkpoint

Set CreateOrderPlex = lOrderPlex

GLogOrder "Create order plex: " & _
            GOrderActionToString(action) & " " & number & " at " & _
            GOrderTypeToString(EntryOrder.OrderType) & " " & lEntryPrice & _
            IIf(lTriggerPrice <> 0, "(" & lTriggerPrice & ")", "") & _
            ", stop " & IIf(StopOrderType = StopOrderTypeAuto, "AUTO(" & stopOffset * mTickSize & ")", lStopPrice) & _
            ", target " & IIf(TargetOrderType = TargetOrderTypeAuto, "AUTO(" & targetOffset * mTickSize & ")", lTargetPrice) & _
            "(Curr price=" & mTicker.FormatPrice(mTicker.TradePrice) & _
            "; id=" & lOrderPlex.Key & ")", _
            mIsSimulated, _
            Me

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

Public Function CreateSellOrderPlex( _
                ByRef number As Long, _
                ByVal EntryOrderType As EntryOrderTypes, _
                ByVal buyPrice As Double, _
                ByVal buyOffset As Long, _
                ByVal TriggerPrice As Double, _
                ByVal StopOrderType As StopOrderTypes, _
                ByVal stopPrice As Double, _
                ByVal stopOffset As Long, _
                ByVal stopLimitPrice As Long, _
                ByVal TargetOrderType As TargetOrderTypes, _
                ByVal targetPrice As Double, _
                ByVal targetOffset As Long, _
                ByVal targetTriggerPrice As Double) As orderPlex
Const ProcName As String = "CreateSellOrderPlex"
Dim failpoint As Long
On Error GoTo Err

Set CreateSellOrderPlex = CreateOrderPlex(OrderActions.ActionSell, _
                                    number, _
                                    EntryOrderType, _
                                    buyPrice, _
                                    buyOffset, _
                                    TriggerPrice, _
                                    StopOrderType, _
                                    stopPrice, _
                                    stopOffset, _
                                    stopLimitPrice, _
                                    TargetOrderType, _
                                    targetPrice, _
                                    targetOffset, _
                                    targetTriggerPrice)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub ExecuteOrderPlex(ByVal pOrderPlex As orderPlex, _
                    Optional checkMoneyManagement As Boolean, _
                    Optional cancelPrice As Double, _
                    Optional StopStrategy As StopManagementStrategy, _
                    Optional ExitStrategy As PositionManagementStrategy, _
                    Optional ScalingStrategy As PositionManagementStrategy)
                    
Dim msg As String
Dim origQuantity As Long
Dim NewQuantity As Long

Const ProcName As String = "ExecuteOrderPlex"
Dim failpoint As Long
On Error GoTo Err

If pOrderPlex Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            "OrderContext::executeOrderPlex", _
                            "No order plex supplied"

If checkMoneyManagement And Not mIsSimulated Then
    origQuantity = pOrderPlex.EntryOrder.Quantity
    NewQuantity = origQuantity
    If Not mPositionManager.AllowOrder( _
                            IIf(pOrderPlex.EntryOrder.action = ActionBuy, True, False), _
                            NewQuantity, _
                            pOrderPlex.Risk, _
                            msg) _
    Then
        GLogOrder "Rejected by money management: " & msg & " " & _
                            "(Curr price=" & mTicker.FormatPrice(mTicker.TradePrice) & _
                            "; id=" & pOrderPlex.Key & ")", _
                            mIsSimulated, _
                            Me
        Set pOrderPlex = Nothing
        Exit Sub
    End If
    If NewQuantity <> origQuantity Then
        pOrderPlex.NewQuantity = NewQuantity
        pOrderPlex.Update
    End If
End If

If Not StopStrategy Is Nothing Then
    StopStrategy.Ticker = mTicker
    StopStrategy.orderPlex = pOrderPlex
    pOrderPlex.StopStrategy = StopStrategy
End If

If Not ExitStrategy Is Nothing Then
    ExitStrategy.Ticker = mTicker
    ExitStrategy.orderPlex = pOrderPlex
    ExitStrategy.OrderContext = Me
    ExitStrategy.OrderContext = Me
End If

If Not ScalingStrategy Is Nothing Then
    ScalingStrategy.Ticker = mTicker
    ScalingStrategy.orderPlex = pOrderPlex
    ScalingStrategy.OrderContext = Me
    pOrderPlex.ScalingStrategy = ScalingStrategy
End If

GLogOrder "Execute order plex " & _
            "(Curr price=" & mTicker.FormatPrice(mTicker.TradePrice) & _
            "; id=" & pOrderPlex.Key & ")", _
            mIsSimulated, _
            Me

mActiveOrderPlexes.Add pOrderPlex
pOrderPlex.Execute cancelPrice

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Function Finish()
Dim op As orderPlex

Const ProcName As String = "Finish"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = Nothing
Set mPositionManager = Nothing
For Each op In mActiveOrderPlexes
    op.Finish
Next
Set mActiveOrderPlexes = Nothing

For Each op In mClosedOrderPlexes
    op.Finish
Next
Set mClosedOrderPlexes = Nothing

Set mOrderSubmitter = Nothing

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetAccountBalance() As Double
Const ProcName As String = "GetAccountBalance"
Dim failpoint As Long
On Error GoTo Err

GetAccountBalance = 10000                       '!!!!!!!!!!!!!!!!!!!!!!!!! need to fix this up

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function IsAttributeSupported( _
                ByVal attributeId As OrderAttributeIds) As Boolean
IsAttributeSupported = (mPermittedOrderAttributes And attributeId)
End Function

Public Function IsOrderTifValueSupported( _
                ByVal value As OrderTifs) As Boolean
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
Dim failpoint As Long
On Error GoTo Err

Set lOrder = New Order
With lOrder
    .Id = GenerateGUIDString
    .IsSimulated = mIsSimulated
    .OrderContext = Me
    .Ticker = mTicker
    If IsAttributeSupported(OrderAttributeIds.OrderAttAllOrNone) Then .AllOrNone = mAllOrNone
    If IsAttributeSupported(OrderAttributeIds.OrderAttBlockOrder) Then .BlockOrder = mBlockOrder
    If IsAttributeSupported(OrderAttributeIds.OrderAttDiscretionaryAmount) Then .DiscretionaryAmount = mDiscretionaryAmount
    If IsAttributeSupported(OrderAttributeIds.OrderAttDisplaySize) Then .DisplaySize = mDisplaySize
    If IsAttributeSupported(OrderAttributeIds.OrderAttETradeOnly) Then .ETradeOnly = mETradeOnly
    If IsAttributeSupported(OrderAttributeIds.OrderAttFirmQuoteOnly) Then .FirmQuoteOnly = mFirmQuoteOnly
    If IsAttributeSupported(OrderAttributeIds.OrderAttGoodAfterTime) Then .GoodAfterTime = mGoodAfterTime
    If IsAttributeSupported(OrderAttributeIds.OrderAttGoodAfterTimeTZ) Then .GoodAfterTimeTZ = mGoodAfterTimeTZ
    If IsAttributeSupported(OrderAttributeIds.OrderAttGoodTillDate) Then .GoodTillDate = mGoodTillDate
    If IsAttributeSupported(OrderAttributeIds.OrderAttGoodTillDateTZ) Then .GoodTillDateTZ = mGoodTillDateTZ
    If IsAttributeSupported(OrderAttributeIds.OrderAttHidden) Then .Hidden = mHidden
    If IsAttributeSupported(OrderAttributeIds.OrderAttIgnoreRTH) Then .IgnoreRegularTradingHours = mIgnoreRegularTradingHours
    If IsAttributeSupported(OrderAttributeIds.OrderAttLimitPrice) Then .LimitPrice = mLimitPrice
    If IsAttributeSupported(OrderAttributeIds.OrderAttMinimumQuantity) Then .MinimumQuantity = mMinimumQuantity
    If IsAttributeSupported(OrderAttributeIds.OrderAttNBBOPriceCap) Then .NbboPriceCap = mNBBOPriceCap
    If mOrderType <> OrderTypes.OrderTypeNone Then
        If IsAttributeSupported(OrderAttributeIds.OrderAttOrderType) Then .OrderType = mOrderType
    End If
    If IsAttributeSupported(OrderAttributeIds.OrderAttOrigin) Then .Origin = mOrigin
    If IsAttributeSupported(OrderAttributeIds.OrderAttOverrideConstraints) Then .OverrideConstraints = mOverrideConstraints
    If IsAttributeSupported(OrderAttributeIds.OrderAttPercentOffset) Then .PercentOffset = mPercentOffset
    If IsAttributeSupported(OrderAttributeIds.OrderAttQuantity) Then .Quantity = mQuantity
    If IsAttributeSupported(OrderAttributeIds.OrderAttSettlingFirm) Then .SettlingFirm = mSettlingFirm
    
    If mStopTriggerMethod <> StopTriggerMethods.StopTriggerNone Then
        If IsAttributeSupported(OrderAttributeIds.OrderAttStopTriggerMethod) Then .StopTriggerMethod = mStopTriggerMethod
    End If
    If IsAttributeSupported(OrderAttributeIds.OrderAttSweepToFill) Then .SweepToFill = mSweepToFill
    If mTimeInForce <> OrderTifs.TIFNone Then
        If IsAttributeSupported(OrderAttributeIds.OrderAttTimeInForce) Then .TimeInForce = mTimeInForce
    End If
    If IsAttributeSupported(OrderAttributeIds.OrderAttTriggerPrice) Then .TriggerPrice = mTriggerPrice
End With

mPositionManager.Orders.Add lOrder

Set NewOrder = lOrder

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Sub NotifyEndOfTradingSession( _
                ByVal pTimestamp As Date)
Const ProcName As String = "NotifyEndOfTradingSession"
Dim failpoint As String
On Error GoTo Err

Dim op As orderPlex

For Each op In mActiveOrderPlexes
    op.NotifyEndOfTradingSession pTimestamp
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub NotifyPositionCloseout( _
                ByVal pPrice As Double, _
                ByVal pTimestamp As Date)
Const ProcName As String = "NotifyCloseout"
Dim failpoint As String
On Error GoTo Err

Dim op As orderPlex

For Each op In mClosedOrderPlexes
    op.NotifyCloseout pPrice, pTimestamp
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Function orderPlexClosed( _
                ByVal pOrderPlex As orderPlex)
Const ProcName As String = "orderPlexClosed"
Dim failpoint As Long
On Error GoTo Err

mClosedOrderPlexes.Add pOrderPlex
mActiveOrderPlexes.Remove pOrderPlex
mCompletedSize = mCompletedSize + pOrderPlex.size
' now check if there are any active order plexes left -
' if not, need to tell the Position Manager
If mActiveOrderPlexes.Count = 0 Then
    mPositionManager.NotifyOrderContextInactive Me
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function Sell(ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal sellPrice As Double, _
                    ByVal sellOffset As Long, _
                    ByVal TriggerPrice As Double, _
                    ByVal StopOrderType As StopOrderTypes, _
                    ByVal stopPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Long, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    Optional ByVal cancelPrice As Double, _
                    Optional ByVal entryTIF As OrderTifs = TIFDay, _
                    Optional ByVal stopTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal targetTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal StopStrategy As StopManagementStrategy, _
                    Optional ByVal ExitStrategy As PositionManagementStrategy, _
                    Optional ByVal ScalingStrategy As PositionManagementStrategy) As orderPlex

Const ProcName As String = "Sell"
Dim failpoint As Long
On Error GoTo Err

GLog pMsg:="Sell", _
    pMsgQualifier:=mTicker.Contract.specifier.LocalSymbol & _
                " number=" & number & _
                "; EntryOrderType=" & GEntryOrderTypeToShortString(EntryOrderType) & _
                "; sellPrice=" & sellPrice & _
                "; sellOffset=" & sellOffset & _
                "; triggerPrice=" & TriggerPrice & _
                "; StopOrderType=" & GStopOrderTypeToShortString(StopOrderType) & _
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
    pProcName:="Sell", pProjName:=ProjectName, pModName:=ModuleName
                
Set Sell = CreateOrderPlex(OrderActions.ActionSell, _
                                number, _
                                EntryOrderType, _
                                sellPrice, _
                                sellOffset, _
                                TriggerPrice, _
                                StopOrderType, _
                                stopPrice, _
                                stopOffset, _
                                stopLimitPrice, _
                                TargetOrderType, _
                                targetPrice, _
                                targetOffset, _
                                targetTriggerPrice)

Sell.EntryOrder.TimeInForce = entryTIF
If Not Sell.stoporder Is Nothing Then Sell.stoporder.TimeInForce = stopTIF
If Not Sell.TargetOrder Is Nothing Then Sell.TargetOrder.TimeInForce = targetTIF

ExecuteOrderPlex Sell, True, cancelPrice, StopStrategy, ExitStrategy, ScalingStrategy

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attributeValueNotSupported( _
                ByVal pAttribute As OrderAttributeIds, _
                ByVal value As Variant)
Const ProcName As String = "attributeValueNotSupported"
Dim failpoint As Long
On Error GoTo Err

Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Value " & CStr(value) & " for attribute " & GOrderAttributeToString(pAttribute) & " is not supported by the service provider"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub checkAttributeSupported( _
                ByVal pAttribute As OrderAttributeIds)
Const ProcName As String = "checkAttributeSupported"
Dim failpoint As Long
On Error GoTo Err

If Not IsAttributeSupported(pAttribute) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                , _
                "The " & GOrderAttributeToString(pAttribute) & " attribute is not supported by the service provider"
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub checkReady()
Const ProcName As String = "checkReady"
Dim failpoint As Long
On Error GoTo Err

If IsReady Then
    mPermittedOrderAttributes = mOrderSubmitter.permittedOrderAttributes
    mPermittedOrderTifs = mOrderSubmitter.PermittedOrderTifs
    mPermittedOrderTypes = mOrderSubmitter.permittedOrderTypes
    mPermittedStopTriggerMethods = mOrderSubmitter.permittedStopTriggerMethods
    RaiseEvent ready
Else
    RaiseEvent NotReady
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function nextPlexKey() As String
Static currID As Long
Const ProcName As String = "nextPlexKey"
Dim failpoint As Long
On Error GoTo Err

currID = currID + 1
nextPlexKey = gTB.SessionID & "/" & currID

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

