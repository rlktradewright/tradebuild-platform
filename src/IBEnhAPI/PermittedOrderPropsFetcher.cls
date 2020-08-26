VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PermittedOrderPropsFetcher"
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

Private Const ModuleName                            As String = "PermittedOrderPropsFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mContractRequester                          As ContractsTwsRequester
Private mContractSpec                               As IContractSpecifier

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mFutureBuilder                              As New FutureBuilder

Private mContractCache                              As ContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
ElseIf ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf TypeOf ev.Future.Value Is IContract Then
    Dim lContract As IContract
    Set lContract = ev.Future.Value
    Set mContractSpec = lContract.Specifier
    mFutureWaiter.Add gFetchContracts(mContractRequester, _
                                        mContractCache, _
                                        mContractSpec, _
                                        Nothing, _
                                        ev.Future.Cookie, _
                                        True)
ElseIf TypeOf ev.Future.Value Is EnumerableCollection Then
    Dim lContracts As EnumerableCollection
    Set lContracts = ev.Future.Value
    Assert Not lContracts Is Nothing, "Failed to retrieve TWS contract details"
    Assert lContracts.Count = 1, "Contract not uniquely specified"
    processTwsContract lContracts.Item(1)
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
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
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

Private Sub processTwsContract( _
                ByVal pTwsContract As TwsContract)
Const ProcName As String = "processTwsContract"
On Error GoTo Err

Dim lOrderTypesStr() As String
lOrderTypesStr = Split(pTwsContract.OrderTypes, ",")

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

mFutureBuilder.Value = lPermittedOrderProps
mFutureBuilder.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




