VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PermittedOrderPropsFetchTask"
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

Implements IContractDetailsConsumer
Implements ITask

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

Private Const ModuleName                            As String = "PermittedOrderPropsFetchTask"

'@================================================================================
' Member variables
'@================================================================================

Private mContractRequester                          As ContractDetailsRequester
Private mContractSpec                               As IContractSpecifier

Private mContractDetails                            As TwsContractDetails

Private mTaskContext                                As TaskContext

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mFutureBuilder                              As New FutureBuilder

Private mContractCache                              As TwsContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'================================================================================
' IContractDetailsConsumer Interface Members
'================================================================================

Private Sub IContractDetailsConsumer_EndContractDetails(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContractDetails"
On Error GoTo Err

processEnd

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_EndOptionParameters(ByVal pRequestIs As Long)

End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, pContractDetails As TwsContractDetails)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

Assert Not pContractDetails Is Nothing, "Contract not uniquely specified"
Set mContractDetails = pContractDetails

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

mTaskContext.Error pErrorCode, pErrorMsg
mFutureBuilder.Fail pErrorCode, pErrorMsg, ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyMarketRule(ByVal pMarketRuleId As Long, pPriceIncrement() As IBAPIV100.TwsPriceIncrement)

End Sub

Private Sub IContractDetailsConsumer_NotifyOptionParameter(ByVal pRequestId As Long, ByVal pExchange As String, ByVal pUnderlyingContractId As Long, ByVal pTradingClass As String, ByVal pMultiplier As Long, pExpirations() As String, pStrikes() As Double)

End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

If mContractSpec Is Nothing Then
    mTaskContext.Suspend -1
    Exit Sub
End If

Set mContractDetails = mContractCache.Item(mContractSpec)

If Not mContractDetails Is Nothing Then
    gLogger.Log "TwsContractDetails retrieved from cache for IContract: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelDetail
    processEnd
    Exit Sub
End If

mContractRequester.RequestContractDetails mContractSpec, Me
mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    mTaskContext.Finish Empty, True
    mFutureBuilder.Cancel
ElseIf ev.Future.IsFaulted Then
    mTaskContext.Error ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
Else
    Dim lContract As IContract
    Set lContract = ev.Future.Value
    Set mContractSpec = lContract.Specifier
    mTaskContext.Continue
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Future() As IFuture
Set Future = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As TwsContractCache, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

mFutureWaiter.Add pContractFuture
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
mFutureBuilder.Cookie = pCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processEnd()
Const ProcName As String = "processEnd"
On Error GoTo Err

Assert Not mContractDetails Is Nothing, "Failed to retrieve contract details"

Dim lOrderTypesStr() As String
lOrderTypesStr = Split(mContractDetails.OrderTypes, ",")

Dim lPermittedOrderAttributes As OrderAttributes
lPermittedOrderAttributes = OrderAttributes.OrderAttBlockOrder
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttIgnoreRTH
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttMinimumQuantity
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttNBBOPriceCap
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttOriginatorRef
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttOverrideConstraints
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttPercentOffset
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttRule80A
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttAction
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttLimitPrice
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttOrderType
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttQuantity
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttTimeInForce
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttTriggerPrice
lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttStopTriggerMethod

Dim lPermittedOrderTypes As OrderTypes
Dim lPermittedOrderTifs As OrderTIFs
Dim i As Long
For i = 0 To UBound(lOrderTypesStr)
    Dim lOrderType As TwsOrderTypes
    lOrderType = gTwsOrderTypeToOrderType(TwsOrderTypeFromString(lOrderTypesStr(i)))
    
    If lOrderType <> OrderTypeNone Then
        lPermittedOrderTypes = lPermittedOrderTypes Or lOrderType
    Else
        Select Case lOrderTypesStr(i)
        Case "AON"  ' all or none
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttAllOrNone
        Case "DAY"  ' today only
            lPermittedOrderTifs = lPermittedOrderTifs Or OrderTIFs.OrderTIFDay
        Case "DIS"  ' discretionary amount
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttDiscretionaryAmount
        Case "ELECONLY"  ' electronic trade only
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttETradeOnly
        Case "FIRMONLY"  ' firm quote only
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttFirmQuoteOnly
        Case "GAT"  ' good after Time
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttGoodAfterTime
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttGoodAfterTimeTZ
        Case "GTC"  ' good till cancelled
            lPermittedOrderTifs = lPermittedOrderTifs Or OrderTIFs.OrderTIFGoodTillCancelled
        Case "GTD"  ' good till date
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDate
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDateTZ
        Case "GTT"  ' good till Time
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDate
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDateTZ
        Case "HID"  ' hidden?
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttHidden
        Case "ICE"  ' iceberg
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttDisplaySize
        Case "IOC"  ' immediate or cancel
            lPermittedOrderTifs = lPermittedOrderTifs Or OrderTIFs.OrderTIFImmediateOrCancel
        Case "RTH"
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttIgnoreRTH
        Case "SWEEP"
            lPermittedOrderAttributes = lPermittedOrderAttributes Or OrderAttributes.OrderAttSweepToFill
        Case "ACTIVETIM", _
            "AD", _
            "ADJUST", _
            "ALERT", _
            "ALLOC", _
            "ALGO", _
            "ALGOCLS", _
            "AVGCOST", _
            "AUC", _
            "BASKET", _
            "BENCHPX", _
            "COND", _
            "CONDORDER", _
            "CONSCOST", _
            "DARKONLY", _
            "DARKPOLL", _
            "DEACT", _
            "DEACTDIS", _
            "DEACTEOD", _
            "FOK", _
            "HPENNY"
        Case "IBDARK", _
            "IBKRATS", _
            "IMB", _
            "LARGE", _
            "LTH", _
            "MIDPX", _
            "NGCOMB"
        Case "NODARK", _
            "NONALGO", _
            "OCA", _
            "OPENCLOSE", _
            "OPG", _
            "OPGREROUT", _
            "PAON", _
            "PEGBENCH", _
            "PEGMID", _
            "PEGMIDVOL", _
            "PEGMKTVOL", _
            "PEGPRMVOL", _
            "PEGSRFVOL", _
            "POSTONLY", _
            "PREOPGRTH", _
            "RELSTK", _
            "RPI", _
            "RTHIGNOPG", _
            "SCALE", _
            "SCALEODD", _
            "SCALERST", _
            "SMARTSTG"
        Case "SNAPMID", _
            "SNAPMKT", _
            "SNAPREL", _
            "SOI", _
            "TIMEPRIO", _
            "TRAILLIT", _
            "TRAILMIT", _
            "VOLAT", _
            "WHATIF"

            ' not sure what these signify
        Case Else
            If lOrderTypesStr(i) <> "" Then
                gLog "Unknown Order type received from Tws", ModuleName, ProcName, pMsgQualifier:=lOrderTypesStr(i)
            End If
        End Select
    End If
Next

Dim lPermittedStopTriggerMethods As OrderStopTriggerMethods
lPermittedStopTriggerMethods = OrderStopTriggerMethods.OrderStopTriggerBidAsk Or _
                                OrderStopTriggerMethods.OrderStopTriggerDefault Or _
                                OrderStopTriggerMethods.OrderStopTriggerDoubleBidAsk Or _
                                OrderStopTriggerMethods.OrderStopTriggerDoubleLast Or _
                                OrderStopTriggerMethods.OrderStopTriggerLast Or _
                                OrderStopTriggerMethods.OrderStopTriggerLastOrBidAsk Or _
                                OrderStopTriggerMethods.OrderStopTriggerMidPoint

Dim lPermittedOrderProps As New PermittedOrderProperties
lPermittedOrderProps.PermittedOrderAttributes = lPermittedOrderAttributes
lPermittedOrderProps.PermittedOrderTifs = lPermittedOrderTifs
lPermittedOrderProps.PermittedOrderTypes = lPermittedOrderTypes
lPermittedOrderProps.PermittedStopTriggerMethods = lPermittedStopTriggerMethods

mTaskContext.Finish lPermittedOrderProps, False
mFutureBuilder.Value = lPermittedOrderProps
mFutureBuilder.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




