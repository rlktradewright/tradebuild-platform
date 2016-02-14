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
                    ByVal OrderType As EntryOrderTypes, _
                    Optional ByVal Price As Double, _
                    Optional ByVal Offset As Long, _
                    Optional ByVal TriggerPrice As Double, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFDay) As ResourceIdentifier
Const ProcName As String = "DeclareEntryOrder"
On Error GoTo Err

Assert Not gTradingContext Is Nothing, "Method can only be called during strategy execution"

Select Case OrderType
Case EntryOrderTypeMarket, _
        EntryOrderTypeMarketOnOpen, _
        EntryOrderTypeMarketOnClose, _
        EntryOrderTypeMarketToLimit
    AssertArgument Price = 0#, "Price must be 0.0 for this order type"
    AssertArgument Offset = 0, "Offset must be 0 for this order type"
    AssertArgument TriggerPrice = 0#, "TriggerPrice must be 0.0 for this order type"
Case EntryOrderTypeMarketIfTouched
    AssertArgument Price = 0#, "Price must be 0.0 for this order type"
    AssertArgument Offset = 0, "Offset must be 0 for this order type"
    AssertArgument TriggerPrice <> 0#, "TriggerPrice must not be 0.0 for this order type"
Case EntryOrderTypeBid, _
        EntryOrderTypeAsk, _
        EntryOrderTypeLast
    AssertArgument Price = 0#, "Price must be 0.0 for this order type"
    AssertArgument TriggerPrice = 0#, "TriggerPrice must be 0.0 for this order type"
Case EntryOrderTypeLimit, _
        EntryOrderTypeLimitOnOpen, _
        EntryOrderTypeLimitOnClose
    AssertArgument TriggerPrice = 0#, "TriggerPrice must be 0.0 for this order type"
Case EntryOrderTypeLimitIfTouched, _
        EntryOrderTypeStopLimit
    AssertArgument Price <> 0#, "Price must not be 0.0 for this order type"
    AssertArgument TriggerPrice <> 0#, "TriggerPrice must not be 0.0 for this order type"
Case EntryOrderTypeStop
    AssertArgument Price = 0#, "Price must be 0.0 for this order type"
    AssertArgument Offset = 0, "Offset must be 0 for this order type"
    AssertArgument TriggerPrice <> 0#, "TriggerPrice must not be 0.0 for this order type"
Case Else
    AssertArgument False, "Invalid entry order type"
End Select

Select Case TimeInForce
Case OrderTIFDay
Case OrderTIFGoodTillCancelled
Case OrderTIFImmediateOrCancel
Case Else
    AssertArgument False, "Invalid TimeInForce"
End Select

Dim os As New OrderSpecifier
os.Initialise OrderRoles.OrderRoleEntry, OrderType, Price, Offset, TriggerPrice, TimeInForce

Set DeclareEntryOrder = gCreateResourceIdentifier(os)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareStopLossOrder( _
                    ByVal OrderType As StopLossOrderTypes, _
                    Optional ByVal Price As Double, _
                    Optional ByVal Offset As Long, _
                    Optional ByVal TriggerPrice As Long, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFGoodTillCancelled) As ResourceIdentifier
Const ProcName As String = "DeclareStopLossOrder"
On Error GoTo Err

Assert Not gTradingContext Is Nothing, "Method can only be called during strategy execution"

Select Case OrderType
Case StopLossOrderTypeStop
    AssertArgument Price = 0#, "Price must be 0.0 for this order type"
    AssertArgument Offset = 0, "Offset must be 0 for this order type"
    AssertArgument TriggerPrice <> 0#, "TriggerPrice must not be 0.0 for this order type"
Case StopLossOrderTypeStopLimit
    AssertArgument Price <> 0#, "Price must not be 0.0 for this order type"
    AssertArgument TriggerPrice <> 0#, "TriggerPrice must not be 0.0 for this order type"
Case StopLossOrderTypeBid, _
        StopLossOrderTypeAsk, _
        StopLossOrderTypeLast, _
        StopLossOrderTypeAuto
    AssertArgument Price = 0#, "Price must be 0.0 for this order type"
    AssertArgument TriggerPrice = 0#, "TriggerPrice must be 0.0 for this order type"
Case Else
    AssertArgument False, "Invalid stop-loss order type"
End Select

Select Case TimeInForce
Case OrderTIFDay
Case OrderTIFGoodTillCancelled
Case OrderTIFImmediateOrCancel
Case Else
    AssertArgument False, "Invalid TimeInForce"
End Select

Dim os As New OrderSpecifier
os.Initialise OrderRoles.OrderRoleStopLoss, OrderType, Price, Offset, TriggerPrice, TimeInForce

Set DeclareStopLossOrder = gCreateResourceIdentifier(os)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareTargetOrder( _
                    ByVal OrderType As TargetOrderTypes, _
                    Optional ByVal Price As Double, _
                    Optional ByVal Offset As Long, _
                    Optional ByVal TriggerPrice As Long, _
                    Optional ByVal TimeInForce As OrderTIFs = OrderTIFGoodTillCancelled) As ResourceIdentifier
Const ProcName As String = "DeclareTargetOrder"
On Error GoTo Err

Assert Not gTradingContext Is Nothing, "Method can only be called during strategy execution"

Select Case OrderType
Case TargetOrderTypeLimit
    AssertArgument TriggerPrice = 0#, "TriggerPrice must be 0.0 for this order type"
Case TargetOrderTypeLimitIfTouched
    AssertArgument Price <> 0#, "Price must not be 0.0 for this order type"
    AssertArgument TriggerPrice <> 0#, "TriggerPrice must not be 0.0 for this order type"
Case TargetOrderTypeMarketIfTouched
    AssertArgument Price = 0#, "Price must be 0.0 for this order type"
    AssertArgument Offset = 0, "Offset must be 0 for this order type"
    AssertArgument TriggerPrice <> 0#, "TriggerPrice must not be 0.0 for this order type"
Case TargetOrderTypeBid, _
        TargetOrderTypeAsk, _
        TargetOrderTypeLast, _
        TargetOrderTypeAuto
    AssertArgument Price = 0#, "Price must be 0.0 for this order type"
    AssertArgument TriggerPrice = 0#, "TriggerPrice must be 0.0 for this order type"
Case Else
    AssertArgument False, "Invalid target order type"
End Select

Select Case TimeInForce
Case OrderTIFDay
Case OrderTIFGoodTillCancelled
Case OrderTIFImmediateOrCancel
Case Else
    AssertArgument False, "Invalid TimeInForce"
End Select

Dim os As New OrderSpecifier
os.Initialise OrderRoles.OrderRoleStopLoss, OrderType, Price, Offset, TriggerPrice, TimeInForce

Set DeclareTargetOrder = gCreateResourceIdentifier(os)

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

Set DeclareOrderContext = gCreateResourceIdentifier(oc)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DeclareProtectedOrderContext(ByVal Name As String) As ResourceIdentifier
Const ProcName As String = "DeclareProtectedOrderContext"
On Error GoTo Err

Assert Not gInitialisationContext Is Nothing, "Method can only be called during strategy initialisation"

Dim oc As OrderContext
Set oc = gInitialisationContext.PositionManager.OrderContexts.Add(Name)
oc.PreventUnprotectedPositions = True

Set DeclareProtectedOrderContext = gCreateResourceIdentifier(oc)

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




