VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Declarators"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "Declarators"

'@================================================================================
' Member variables
'@================================================================================

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

'@================================================================================
' Methods
'@================================================================================

Public Function DeclareChart(ByVal Timeframe As String) As ResourceIdentifier

End Function

Public Function DeclareEntryOrder( _
                    ByVal OrderType As OrderTypes, _
                    Optional ByVal LimitPrice As String, _
                    Optional ByVal TriggerPrice As String, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFDay, _
                    Optional ByVal IgnoreRTH As Boolean = False) As ResourceIdentifier
Const ProcName As String = "DeclareEntryOrder"
On Error GoTo Err

Set DeclareEntryOrder = doDeclareOrder( _
                OrderType, _
                OrderRoleEntry, _
                False, _
                LimitPrice, _
                TriggerPrice, _
                TimeInForce, _
                IgnoreRTH)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareSimulatedEntryOrder( _
                    ByVal OrderType As OrderTypes, _
                    Optional ByVal LimitPrice As String, _
                    Optional ByVal TriggerPrice As String, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFDay, _
                    Optional ByVal IgnoreRTH As Boolean = False) As ResourceIdentifier
Const ProcName As String = "DeclareSimulatedEntryOrder"
On Error GoTo Err

Set DeclareSimulatedEntryOrder = doDeclareOrder( _
                OrderType, _
                OrderRoleEntry, _
                True, _
                LimitPrice, _
                TriggerPrice, _
                TimeInForce, _
                IgnoreRTH)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareStopLossOrder( _
                    ByVal OrderType As StopLossOrderTypes, _
                    Optional ByVal LimitPrice As String, _
                    Optional ByVal TriggerPrice As String, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFDay, _
                    Optional ByVal IgnoreRTH As Boolean = False) As ResourceIdentifier
Const ProcName As String = "DeclareStopLossOrder"
On Error GoTo Err

Set DeclareStopLossOrder = doDeclareOrder( _
                OrderType, _
                OrderRoleStopLoss, _
                False, _
                LimitPrice, _
                TriggerPrice, _
                TimeInForce, _
                IgnoreRTH)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareSimulatedStopLossOrder( _
                    ByVal OrderType As StopLossOrderTypes, _
                    Optional ByVal LimitPrice As String, _
                    Optional ByVal TriggerPrice As String, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFDay, _
                    Optional ByVal IgnoreRTH As Boolean = False) As ResourceIdentifier
Const ProcName As String = "DeclareSimulatedStopLossOrder"
On Error GoTo Err

Set DeclareSimulatedStopLossOrder = doDeclareOrder( _
                OrderType, _
                OrderRoleStopLoss, _
                True, _
                LimitPrice, _
                TriggerPrice, _
                TimeInForce, _
                IgnoreRTH)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareTargetOrder( _
                    ByVal OrderType As TargetOrderTypes, _
                    Optional ByVal LimitPrice As String, _
                    Optional ByVal TriggerPrice As String, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFDay, _
                    Optional ByVal IgnoreRTH As Boolean = False) As ResourceIdentifier
Const ProcName As String = "DeclareTargetOrder"
On Error GoTo Err

Set DeclareTargetOrder = doDeclareOrder( _
                OrderType, _
                OrderRoleTarget, _
                False, _
                LimitPrice, _
                TriggerPrice, _
                TimeInForce, _
                IgnoreRTH)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareSimulatedTargetOrder( _
                    ByVal OrderType As TargetOrderTypes, _
                    Optional ByVal LimitPrice As String, _
                    Optional ByVal TriggerPrice As String, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFDay, _
                    Optional ByVal IgnoreRTH As Boolean = False) As ResourceIdentifier
Const ProcName As String = "DeclareSimulatedTargetOrder"
On Error GoTo Err

Set DeclareSimulatedTargetOrder = doDeclareOrder( _
                OrderType, _
                OrderRoleTarget, _
                True, _
                LimitPrice, _
                TriggerPrice, _
                TimeInForce, _
                IgnoreRTH)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareOrderContext(ByVal Name As String) As ResourceIdentifier
Const ProcName As String = "DeclareOrderContext"
On Error GoTo Err

Assert Not gInitialisationContext Is Nothing, "Method can only be called during strategy initialisation"

Dim oc As OrderContext
Set oc = gInitialisationContext.PositionManager.OrderContexts.Add(Name)
If Not gInitialisationContext.AllowUnprotectedPositions Then oc.PreventUnprotectedPositions = True

Set DeclareOrderContext = gCreateResourceIdentifier(oc)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareSimulatedOrderContext(ByVal Name As String) As ResourceIdentifier
Const ProcName As String = "DeclareSimulatedOrderContext"
On Error GoTo Err

Assert Not gInitialisationContext Is Nothing, "Method can only be called during strategy initialisation"

Dim oc As OrderContext
Set oc = gInitialisationContext.PositionManagerSimulated.OrderContexts.Add(Name)

Set DeclareSimulatedOrderContext = gCreateResourceIdentifier(oc)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareStudy(ByVal Name As String, _
                Optional ByVal BasedOn As ResourceIdentifier, _
                Optional ByVal Parameters As String, _
                Optional ByVal IncludeBarsOutsideSession As Boolean = False, _
                Optional ByVal NumberOfValuesToCache As Long = 3, _
                Optional ByVal InputNames As String) As ResourceIdentifier
Const ProcName As String = "DeclareStudy"
On Error GoTo Err

Assert Not gInitialisationContext Is Nothing, "Method can only be called during strategy initialisation"

If InputNames = "" Then InputNames = DefaultStudyValueName

Dim lInputNames() As String
lInputNames = Split(InputNames, ",")

If BasedOn Is Nothing Then Set BasedOn = gResourceContext.PrimaryTimeframe

Dim lBaseStudy As IStudy

Dim lBaseObj As Object
Set lBaseObj = BasedOn.Resource
If TypeOf lBaseObj Is IStudy Then
    Set lBaseStudy = lBaseObj
ElseIf TypeOf lBaseObj Is Timeframe Then
    Dim tf As Timeframe
    Set tf = lBaseObj
    Set lBaseStudy = tf.BarStudy
Else
    AssertArgument False, "Specified resource is not a study or a timeframe"
End If

Dim lLibraryName As String
Dim lStudyName As String
Dim p As Long: p = InStr(1, Name, "\")
If p = 0 Then
    lStudyName = Name
Else
    AssertArgument p <> 1, "Study name cannot start with '\'"
    lLibraryName = Left$(Name, p - 1)
    lStudyName = Right$(Name, Len(Name) - p)
End If

Dim lStudy As IStudy
Set lStudy = gInitialisationContext.AddStudy( _
                    lStudyName, _
                    lBaseStudy, _
                    lInputNames, _
                    IncludeBarsOutsideSession, _
                    CreateParametersFromString(Parameters), _
                    NumberOfValuesToCache, _
                    lLibraryName)

Set DeclareStudy = gCreateResourceIdentifier(lStudy)
gResourceContext.SetPrimaryStudyOfType DeclareStudy

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareTimeframe( _
                ByVal BarLength As Long, _
                Optional ByVal BarUnit As String = "minutes", _
                Optional ByVal NumberOfBars As Long = 500, _
                Optional ByVal IncludeBarsOutsideSession As Boolean = False, _
                Optional ByVal ShowInChart As Boolean = True) As ResourceIdentifier
Const ProcName As String = "DeclareTimeframe"
On Error GoTo Err

Assert Not gInitialisationContext Is Nothing, "Method can only be called during strategy initialisation"

Dim tf As Timeframe
Set tf = gInitialisationContext.AddTimeframe(GetTimePeriod(BarLength, TimePeriodUnitsFromString(BarUnit)), NumberOfBars, IncludeBarsOutsideSession, ShowInChart)

Set DeclareTimeframe = gCreateResourceIdentifier(tf)
If gResourceContext.PrimaryTimeframe Is Nothing Then gResourceContext.PrimaryTimeframe = DeclareTimeframe

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkPriceSpecifiers( _
                ByVal pOrderRole As OrderRoles, _
                ByVal pOrderType As Long, _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pTriggerPriceSpec As PriceSpecifier)
Const ProcName As String = "checkPriceSpecifiers"
On Error GoTo Err

Select Case pOrderRole
Case OrderRoleEntry
    Select Case pOrderType
    Case EntryOrderTypeMarket, _
            EntryOrderTypeMarketOnOpen, _
            EntryOrderTypeMarketOnClose, _
            EntryOrderTypeMarketToLimit
        AssertArgument Not pPriceSpec.IsValid, "Price can't be specified for this order type"
        AssertArgument Not pTriggerPriceSpec.IsValid, "TriggerPrice can't be specified for this order type"
    Case EntryOrderTypeMarketIfTouched
        AssertArgument Not pPriceSpec.IsValid, "Price can't be specified for this order type"
        AssertArgument pTriggerPriceSpec.IsValid, "TriggerPrice must be specified for this order type"
    Case EntryOrderTypeLimit, _
            EntryOrderTypeLimitOnOpen, _
            EntryOrderTypeLimitOnClose
        AssertArgument pPriceSpec.IsValid, "Price must be specified for this order type"
        AssertArgument Not pTriggerPriceSpec.IsValid, "TriggerPrice can't be specified for this order type"
    Case EntryOrderTypeLimitIfTouched, _
            EntryOrderTypeStopLimit, _
            EntryOrderTypeTrailStopLimit
        AssertArgument pPriceSpec.IsValid, "Price must be specified for this order type"
        AssertArgument pTriggerPriceSpec.IsValid, "TriggerPrice must be specified for this order type"
    Case EntryOrderTypeStop, _
            EntryOrderTypeTrailStop
        AssertArgument Not pPriceSpec.IsValid, "Price can't be specified for this order type"
        AssertArgument pTriggerPriceSpec.IsValid, "TriggerPrice must be specified for this order type"
    Case Else
        AssertArgument False, "Invalid entry order type"
    End Select
Case OrderRoleStopLoss
    Select Case pOrderType
    Case StopLossOrderTypeStop, _
            StopLossOrderTypeTrailStop
        AssertArgument Not pPriceSpec.IsValid, "Price can't be specified for this order type"
        AssertArgument pTriggerPriceSpec.IsValid, "TriggerPrice must be specified for this order type"
    Case StopLossOrderTypeStopLimit, _
            StopLossOrderTypeTrailStopLimit
        AssertArgument pPriceSpec.IsValid, "Price must be specified for this order type"
        AssertArgument pTriggerPriceSpec.IsValid, "TriggerPrice must be specified for this order type"
    Case Else
        AssertArgument False, "Invalid stop-loss order type"
    End Select
Case OrderRoleTarget
    Select Case pOrderType
    Case TargetOrderTypeLimit
        AssertArgument pPriceSpec.IsValid, "Price must be specified for this order type"
        AssertArgument Not pTriggerPriceSpec.IsValid, "TriggerPrice can't be specified for this order type"
    Case TargetOrderTypeLimitIfTouched
        AssertArgument pPriceSpec.IsValid, "Price must be specified for this order type"
        AssertArgument pTriggerPriceSpec.IsValid, "TriggerPrice must be specified for this order type"
    Case TargetOrderTypeMarketIfTouched
        AssertArgument Not pPriceSpec.IsValid, "Price can't be specified for this order type"
        AssertArgument pTriggerPriceSpec.IsValid, "TriggerPrice must be specified for this order type"
    Case Else
        AssertArgument False, "Invalid target order type"
    End Select
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function doDeclareOrder( _
                ByVal pOrderType As OrderTypes, _
                ByVal pOrderRole As OrderRoles, _
                ByVal pIsSimulated As Boolean, _
                ByVal pLimitPrice As String, _
                ByVal pTriggerPrice As String, _
                ByVal pTimeInForce As OrderTIFs, _
                ByVal pIgnoreRTH As Boolean) As ResourceIdentifier
Const ProcName As String = "doDeclareOrder"
On Error GoTo Err

Assert Not gTradingContext Is Nothing, "Method can only be called during strategy execution"

Select Case pTimeInForce
Case OrderTIFDay
Case OrderTIFGoodTillCancelled
Case OrderTIFImmediateOrCancel
Case Else
    AssertArgument False, "Invalid TimeInForce"
End Select

' any order context will do for our purposes here
Dim lOrderContext As OrderContext
If pIsSimulated Then
    Set lOrderContext = gTradingContext.DefaultBuyOrderContextSimulated
Else
    Set lOrderContext = gTradingContext.DefaultBuyOrderContext
End If

Dim lPriceSpec As PriceSpecifier
Dim lMessage As String
Assert lOrderContext.ParsePriceAndOffset(lPriceSpec, pLimitPrice, lMessage), "Invalid limit price string: " & lMessage
Dim lTriggerPriceSpec As PriceSpecifier
Assert lOrderContext.ParsePriceAndOffset(lTriggerPriceSpec, pTriggerPrice, lMessage), "Invalid trigger price string" & lMessage

checkPriceSpecifiers pOrderRole, pOrderType, lPriceSpec, lTriggerPriceSpec

Dim os As New OrderSpecifier
os.Initialise pOrderRole, pOrderType, pIsSimulated, lPriceSpec, lTriggerPriceSpec, pTimeInForce, pIgnoreRTH

Set doDeclareOrder = gCreateResourceIdentifier(os)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




