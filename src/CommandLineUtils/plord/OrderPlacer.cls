VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPlacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "OrderPlacer"

'@================================================================================
' Member variables
'@================================================================================

Private mBracketOrderSpec                           As BracketOrderSpecification
Private mContract                                   As IContract
Private mOrderContext                               As OrderContext
Attribute mOrderContext.VB_VarHelpID = -1
Private WithEvents mOrderContextEv                  As OrderContext
Attribute mOrderContextEv.VB_VarHelpID = -1
Private mStageOrders                                As Boolean

Private mTimezone                                   As TimeZone

Private mMoneyManager                               As MoneyManager

Private mGroupResources                             As GroupResources

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mOrderContext Event Handlers
'@================================================================================

Private Sub mOrderContextEv_Change(ev As ChangeEventData)
Const ProcName As String = "mOrderContextEv_Change"
On Error GoTo Err

Dim lChangeType As OrderContextChangeTypes: lChangeType = ev.ChangeType

Select Case lChangeType
Case OrderContextChangeTypes.OrderContextActiveStateChanged
    Exit Sub
Case OrderContextChangeTypes.OrderContextReadyStateChanged
    Static sPendingMessageLogged As Boolean
    If Not isReadyForOrder Then
        If Not sPendingMessageLogged Then
            sPendingMessageLogged = True
            gWriteLineToConsole mBracketOrderSpec.ID & ": Order is pending: " & mOrderContextEv.NotReadyMessage, True
        End If
        Exit Sub
    End If
    
    Set mOrderContextEv = Nothing
    gPlaceOrdersTask.AddOrderPlacer Me
    
Case OrderContextChangeTypes.OrderContextFinished
    Cancel mOrderContextEv.NotReadyMessage
End Select


Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Cancel( _
                ByVal pReason As String)
mGroupResources.OrderPlacers.Remove Me
Set mGroupResources = Nothing
gPlaceOrdersTask.RemoveOrderPlacer Me
gWriteLineToConsole mBracketOrderSpec.ID & ": Order cancelled: " & pReason
End Sub

Friend Sub Initialise( _
                ByVal pBracketOrderSpec As BracketOrderSpecification, _
                ByVal pcontract As IContract, _
                ByVal pOrderContext As OrderContext, _
                ByVal pStageOrders As Boolean, _
                ByVal pMoneyManager As MoneyManager, _
                ByVal pGroupResources As GroupResources)
Set mBracketOrderSpec = pBracketOrderSpec
Set mContract = pcontract
Set mOrderContext = pOrderContext
mStageOrders = pStageOrders
Set mMoneyManager = pMoneyManager
Set mGroupResources = pGroupResources

Set mTimezone = GetTimeZone(mContract.TimezoneName)
End Sub

Friend Sub Finish()
Set mOrderContext = Nothing
Set mOrderContextEv = Nothing
End Sub

Friend Function PlaceOrder() As Boolean
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Static sAttempts As Long

Dim lPrefix As String: lPrefix = FormatTimestamp(GetTimestamp, TimestampTimeOnlyISO8601) & " " & _
                                mBracketOrderSpec.ID & ": "

If sAttempts = 0 Then
    LogMessage _
                lPrefix & "TIME " & _
                FormatTimestamp(mTimezone.ConvertDateLocalToTz(GetTimestamp), _
                                TimestampDateAndTimeISO8601)
    LogMessage lPrefix & contractSpecToString
    LogMessage lPrefix & mBracketOrderSpec.ToString
End If
sAttempts = sAttempts + 1

If Not isReadyForOrder Then
    Set mOrderContextEv = mOrderContext
    gWriteLineToConsole lPrefix & "Order is pending: " & mOrderContext.NotReadyMessage
    Exit Function
End If

If sAttempts > 1 Then
    gWriteLineToConsole lPrefix & "Provider is now ready"
    LogMessage _
                lPrefix & "TIME " & _
                FormatTimestamp(mTimezone.ConvertDateLocalToTz(GetTimestamp), _
                                TimestampDateAndTimeISO8601)
    LogMessage lPrefix & contractSpecToString
    LogMessage lPrefix & mBracketOrderSpec.ToString
End If

mOrderContext.OrderSubmitter.OrdersAreStaged = mStageOrders

Dim lBO As BracketOrder

With mBracketOrderSpec
    Dim lEntryOrder As IOrder
    Set lEntryOrder = mOrderContext.CreateEntryOrder( _
                            .EntryType, _
                            .EntryLimitPriceSpec, _
                            .EntryTriggerPriceSpec, _
                            .EntryIgnoreRTH)
    
    Dim lStopLossOrder As IOrder
    If .StopLossType <> OrderTypeNone Then
        Set lStopLossOrder = mOrderContext.CreateStopLossOrder( _
                                .StopLossType, _
                                .StopLossLimitPriceSpec, _
                                .StopLossTriggerPriceSpec, _
                                .StopLossIgnoreRTH)
    End If

    Dim lTargetOrder As IOrder
    If .TargetType <> OrderTypeNone Then
        Set lTargetOrder = mOrderContext.CreateTargetOrder( _
                                .TargetType, _
                                .TargetLimitPriceSpec, _
                                .TargetTriggerPriceSpec, _
                                .TargetIgnoreRTH)
    End If


    Set lBO = mOrderContext.CreateBracketOrderWithId( _
                    .Action, _
                    .Quantity, _
                    lEntryOrder, _
                    .ID, _
                    lStopLossOrder, _
                    lTargetOrder _
                    )
    
    gBracketOrderListener.Add lBO
    
    lBO.CancelAfter = .CancelAfter
    lBO.CancelPrice = .CancelPrice
    lBO.Description = .Description
    lBO.EntryOrder.TimeInForce = .EntryTIF
    lBO.EntryReason = .EntryReason
    If Not lBO.StopLossOrder Is Nothing Then
        lBO.StopLossOrder.TimeInForce = .StopLossTIF
        lBO.StopReason = .StopLossReason
    End If
    If Not lBO.TargetOrder Is Nothing Then
        lBO.TargetOrder.TimeInForce = .TargetTIF
        lBO.TargetReason = .TargetReason
    End If
    If .GoodAfterTime <> 0 Then
        lBO.EntryOrder.GoodAfterTime = .GoodAfterTime
        lBO.EntryOrder.GoodAfterTimeTZ = .TimezoneName
    End If
    If .GoodTillDate <> 0 Then
        lBO.EntryOrder.GoodTillDate = .GoodTillDate
        lBO.EntryOrder.GoodTillDateTZ = .TimezoneName
    End If
    If Not .RolloverSpec Is Nothing Then lBO.RolloverSpecification = .RolloverSpec
End With

gWriteLineToConsole lPrefix & "bracket order ready to be placed"

If mBracketOrderSpec.AutosizeType <> BracketAutosizeNone Then
    lBO.AutosizeType = mBracketOrderSpec.AutosizeType
    lBO.AutosizeValue = mBracketOrderSpec.AutosizeValue
    lBO.LotSize = mBracketOrderSpec.LotSize
End If
mOrderContext.ExecuteBracketOrder lBO, True
If lBO.State = BracketOrderStateSubmitted Then CreateBracketProfitCalculator lBO, mOrderContext.DataSource

PlaceOrder = True

mGroupResources.OrderPlacers.Remove Me
Set mGroupResources = Nothing
gPlaceOrdersTask.RemoveOrderPlacer Me
Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function isReadyForOrder() As Boolean
isReadyForOrder = False
With mOrderContext
    If Not .IsReady Then
    ElseIf Not .IsReadyTickSize() Then
    ElseIf Not .IsReadyForAutoSize(mBracketOrderSpec.AutosizeType, _
                                    mBracketOrderSpec.Contract.Specifier.CurrencyCode) Then
    ElseIf Not .IsReadyForOrderType(mBracketOrderSpec.EntryType) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.EntryLimitPriceSpec) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.EntryTriggerPriceSpec) Then
    ElseIf Not .IsReadyForOrderType(mBracketOrderSpec.StopLossType) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.StopLossLimitPriceSpec) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.StopLossTriggerPriceSpec) Then
    ElseIf Not .IsReadyForOrderType(mBracketOrderSpec.TargetType) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.TargetLimitPriceSpec) Then
    ElseIf Not .IsReadyForPriceSpecifier(mBracketOrderSpec.TargetTriggerPriceSpec) Then
    Else
        isReadyForOrder = True
    End If
End With
End Function

Private Function contractSpecToString() As String
Const ProcName As String = "contractSpecToString"
On Error GoTo Err

Dim s As String
s = gCommands.ContractCommand.Name & " "
s = s & gGenerateSwitch(SymbolSwitch, mContract.Specifier.Symbol)
s = s & gGenerateSwitch(LocalSymbolSwitch, mContract.Specifier.LocalSymbol)
s = s & gGenerateSwitch(SecTypeSwitch, SecTypeToShortString(mContract.Specifier.SecType))
s = s & gGenerateSwitch(ExchangeSwitch, mContract.Specifier.Exchange)
If mContract.Specifier.Expiry <> "" Then s = s & gGenerateSwitch(ExpirySwitch, mContract.Specifier.Expiry)
s = s & gGenerateSwitch(CurrencySwitch, mContract.Specifier.CurrencyCode)
If mContract.Specifier.Multiplier <> 1 Then s = s & gGenerateSwitch(MultiplierSwitch, mContract.Specifier.Multiplier)
If mContract.Specifier.Right <> OptNone Then
    s = s & gGenerateSwitch(RightSwitch, OptionRightToString(mContract.Specifier.Right))
    s = s & gGenerateSwitch(StrikeSwitch, mContract.Specifier.Strike)
End If
contractSpecToString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




