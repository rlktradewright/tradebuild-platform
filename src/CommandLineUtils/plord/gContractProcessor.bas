Attribute VB_Name = "gContractProcessor"
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

Private Const ModuleName                            As String = "gContractProcessor"

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

Public Function gParseOrderPrices( _
                ByVal pPrice1 As String, _
                ByVal pPrice2 As String, _
                ByVal pOrderType As OrderTypes, _
                ByRef pLimitPriceSpec As PriceSpecifier, _
                ByRef pTriggerPriceSpec As PriceSpecifier, _
                ByRef pMessage As String) As Boolean
Const ProcName As String = "gParseOrderPrices"
On Error GoTo Err

gParseOrderPrices = True

Set pLimitPriceSpec = NewPriceSpecifier
Set pTriggerPriceSpec = NewPriceSpecifier

Select Case pOrderType
Case OrderTypeMarket, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketOnOpen, _
        OrderTypeMarketToLimit, _
        OrderTypeMidprice
    If notifyUnexpectedPrice(pPrice1, "limit", pMessage) Then gParseOrderPrices = False
    If notifyUnexpectedPrice(pPrice2, "trigger", pMessage) Then gParseOrderPrices = False
Case OrderTypeMarketIfTouched, _
        OrderTypeStop, _
        OrderTypeTrail
    If Not parseOrderPrice(pPrice1, "trigger", pTriggerPriceSpec, pMessage) Then gParseOrderPrices = False
    If notifyUnexpectedPrice(pPrice2, "second", pMessage) Then gParseOrderPrices = False
Case OrderTypeLimit, _
        OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose
    If Not parseOrderPrice(pPrice1, "limit", pLimitPriceSpec, pMessage) Then gParseOrderPrices = False
    If notifyUnexpectedPrice(pPrice2, "trigger", pMessage) Then gParseOrderPrices = False
Case OrderTypeLimitIfTouched, _
        OrderTypeStopLimit, _
        OrderTypeTrailLimit
    If Not parseOrderPrice(pPrice1, "Trigger", pTriggerPriceSpec, pMessage) Then gParseOrderPrices = False
    If Not parseOrderPrice(pPrice2, "Limit", pLimitPriceSpec, pMessage) Then gParseOrderPrices = False
End Select
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gProcessRolloverCommand( _
                ByVal pParams As String, _
                ByVal pSecType As SecurityTypes) As RolloverSpecification
Const ProcName As String = "gProcessRolloverCommand"
On Error GoTo Err

AssertArgument (pSecType = SecTypeFuture Or _
                    pSecType = SecTypeFuturesOption Or _
                    pSecType = SecTypeOption), _
                "Rollover only applies to options, futures and futures options"

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pParams, " ")

AssertArgument lClp.NumberOfArgs = 0, "Only attributes beginning with ""/"" are permitted"

Dim lQuantityMode As RolloverQuantityModes: lQuantityMode = RolloverQuantityModeAsPrevious
Dim lQuantityParameter As BoxedDecimal
Dim lLotSize As Long
If lClp.Switch(QuantitySwitch) Then processRolloverQuantity _
                        lClp.SwitchValue(QuantitySwitch), _
                        lQuantityMode, _
                        lQuantityParameter, _
                        lLotSize

Dim lDays As Long
If lClp.Switch(DaysSwitch) Then lDays = processRolloverDays(lClp.SwitchValue(DaysSwitch))

Dim lTime As Date
If lClp.Switch(TimeSwitch) Then lTime = processRolloverTime(lClp.SwitchValue(TimeSwitch))

Dim lCloseType As OrderTypes: lCloseType = OrderTypeMarket
Dim lCloseLimitPriceSpec As New PriceSpecifier
Dim lCloseTriggerPriceSpec As New PriceSpecifier
Dim lCloseTimeout As Long
If lClp.Switch(CloseSwitch) Then processRolloverClose _
                                        lClp.SwitchValue(CloseSwitch), _
                                        lCloseType, _
                                        lCloseLimitPriceSpec, _
                                        lCloseTriggerPriceSpec, _
                                        lCloseTimeout

Dim lEntryType As OrderTypes
Dim lEntryLimitPriceSpec As New PriceSpecifier
Dim lEntryTriggerPriceSpec As New PriceSpecifier
Dim lEntryTimeout As Long
If lClp.Switch(EntrySwitch) Then processRolloverEntry _
                                        lClp.SwitchValue(EntrySwitch), _
                                        lEntryType, _
                                        lEntryLimitPriceSpec, _
                                        lEntryTriggerPriceSpec, _
                                        lEntryTimeout

Dim lStrikeMode As RolloverStrikeModes
lStrikeMode = RolloverStrikeModeAsPrevious

Dim lStrikeValue As Double: lStrikeValue = 0

Dim lStrikeOperator As OptionStrikeSelectionOperators
lStrikeOperator = OptionStrikeSelectionOperatorNone

Dim lUnderlyingExchangeName As String

If lClp.Switch(StrikeSwitch) Then processRolloverStrike _
                                        lClp.SwitchValue(StrikeSwitch), _
                                        lStrikeMode, _
                                        lStrikeValue, _
                                        lStrikeOperator, _
                                        lUnderlyingExchangeName

If pSecType = SecTypeOption Then
    Set gProcessRolloverCommand = CreateOptionRolloverSpecification( _
                                lDays, _
                                lTime, _
                                OptionStrikeSelectionModeNone, _
                                0#, _
                                OptionStrikeSelectionOperatorNone, _
                                lStrikeMode, _
                                lStrikeValue, _
                                lStrikeOperator, _
                                lQuantityMode, _
                                lQuantityParameter, _
                                lLotSize, _
                                lUnderlyingExchangeName, _
                                lCloseType, _
                                lCloseTimeout, _
                                lCloseLimitPriceSpec, _
                                lCloseTriggerPriceSpec, _
                                lEntryType, _
                                lEntryTimeout, _
                                lEntryLimitPriceSpec, _
                                lEntryTriggerPriceSpec)
                                
Else
    Set gProcessRolloverCommand = CreateRolloverSpecification( _
                                lDays, _
                                lTime, _
                                lCloseType, _
                                lCloseTimeout, _
                                lCloseLimitPriceSpec, _
                                lCloseTriggerPriceSpec, _
                                lEntryType, _
                                lEntryTimeout, _
                                lEntryLimitPriceSpec, _
                                lEntryTriggerPriceSpec)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function notifyUnexpectedPrice( _
                ByVal pPriceString As String, _
                ByVal pPriceName As String, _
                ByRef pMessage As String) As Boolean
If pPriceString <> "" Then
    If pMessage <> "" Then pMessage = pMessage & vbCrLf
    pMessage = pMessage & pPriceName & " price must not be specified for this order type"
    notifyUnexpectedPrice = True
Else
    notifyUnexpectedPrice = False
End If
End Function

Private Function parseOrderPrice( _
                ByVal pPriceStr As String, _
                ByVal pPriceName As String, _
                ByRef pPriceSpec As PriceSpecifier, _
                ByRef pMessage As String) As Boolean
Const ProcName As String = "parseOrderPrice"
On Error GoTo Err

Dim lMessage As String

If pPriceStr = "" Then
    parseOrderPrice = False
    pMessage = pPriceName & " price must be specified for this order type"
ElseIf ParsePriceAndOffset(pPriceSpec, pPriceStr, SecTypeNone, 0#, lMessage) And _
    pPriceSpec.IsValid _
Then
    parseOrderPrice = True
Else
    parseOrderPrice = False
    pMessage = pPriceName & " " & lMessage & ": " & pPriceStr
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub parseOrderSpec( _
                ByVal pValue As String, _
                ByRef pOrderType As OrderTypes, _
                ByRef pLimitPriceSpec As PriceSpecifier, _
                ByRef pTriggerPriceSpec As PriceSpecifier, _
                ByRef pTimeoutSecs As Long)
Const ProcName As String = "parseOrderSpec"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pValue, ";")

Dim lOrderTypeStr As String: lOrderTypeStr = lClp.Arg(0)

If lOrderTypeStr = "" Then
    pOrderType = OrderTypeNone
Else
    pOrderType = OrderTypeFromString(lOrderTypeStr)
End If

'If Not pOrderContext Is Nothing Then AssertArgument (pOrderContext.PermittedOrderTypes And pOrderType) <> 0, _
'                                                    "Invalid order type: " & lOrderTypeStr

Dim lMaxNumberOfArgsRequired As Long
Dim lPrice1Spec As String
Dim lPrice2Spec As String
Dim lTimeoutSpec As String

Select Case pOrderType
Case OrderTypeMarket, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketOnOpen, _
        OrderTypeMarketToLimit, _
        OrderTypeMidprice
    lMaxNumberOfArgsRequired = 1
Case OrderTypeMarketIfTouched, _
        OrderTypeStop, _
        OrderTypeTrail, _
        OrderTypeLimit, _
        OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose, _
        OrderTypeLimitIfTouched, _
        OrderTypeStopLimit, _
        OrderTypeTrailLimit
    lMaxNumberOfArgsRequired = 3
    lPrice1Spec = lClp.Arg(1)
    lTimeoutSpec = lClp.Arg(2)
Case Else
    lMaxNumberOfArgsRequired = 4
    lPrice1Spec = lClp.Arg(1)
    lPrice2Spec = lClp.Arg(2)
    lTimeoutSpec = lClp.Arg(3)
End Select

AssertArgument lClp.NumberOfArgs <= lMaxNumberOfArgsRequired, _
                "Too many elements in order spec"

Dim lMessage As String
AssertArgument gParseOrderPrices(lPrice1Spec, _
                    lPrice2Spec, _
                    pOrderType, _
                    pLimitPriceSpec, _
                    pTriggerPriceSpec, _
                    lMessage), _
                lMessage

pTimeoutSecs = parseOrderTimeoutSecs(lTimeoutSpec)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function parseOrderTimeoutSecs( _
                ByVal pTimeoutSpec As String) As Double
Const SecondsDesignator As String = "S"
Const MinutesDesignator As String = "M"

If pTimeoutSpec = "" Then
    parseOrderTimeoutSecs = 0
    Exit Function
End If

Dim lUnits As String
Dim lTimeout As String
If UCase$(Right$(pTimeoutSpec, 1)) = SecondsDesignator Then
    lUnits = SecondsDesignator
    lTimeout = Left$(pTimeoutSpec, Len(pTimeoutSpec) - 1)
ElseIf UCase$(Right$(pTimeoutSpec, 1)) = MinutesDesignator Then
    lUnits = MinutesDesignator
    lTimeout = Left$(pTimeoutSpec, Len(pTimeoutSpec) - 1)
Else
    AssertArgument False, "Timeout spec invalid - must be a positive integer " & _
                            "followed by ""M"" for minutes or ""S"" for seconds " & _
                            "(not case-sensitive)"
    Exit Function
End If

If IsInteger(lTimeout, 0) Then
    parseOrderTimeoutSecs = CLng(lTimeout)
    If lUnits = MinutesDesignator Then parseOrderTimeoutSecs = parseOrderTimeoutSecs * 60
Else
    AssertArgument False, "Invalid timeout: " & pTimeoutSpec
End If
End Function

Private Sub processRolloverClose( _
                ByVal pValue As String, _
                ByRef pCloseType As OrderTypes, _
                ByRef pCloseLimitPriceSpec As PriceSpecifier, _
                ByRef pCloseTriggerPriceSpec As PriceSpecifier, _
                ByRef pCloseTimeoutSecs As Long)
Const ProcName As String = "processRolloverClose"
On Error GoTo Err

parseOrderSpec pValue, pCloseType, pCloseLimitPriceSpec, pCloseTriggerPriceSpec, pCloseTimeoutSecs

If Not IsEntryOrderType(pCloseType) Then
    pCloseType = OrderTypeNone
    Set pCloseLimitPriceSpec = New PriceSpecifier
    Set pCloseTriggerPriceSpec = New PriceSpecifier
    pCloseTimeoutSecs = 0
    AssertArgument True, "Invalid close order type"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processRolloverDays(ByVal pValue As String) As Long
AssertArgument IsInteger(pValue, 0, 30), _
            "Days must be an integer from 0 to 30"
    
processRolloverDays = CLng(pValue)
End Function

Private Sub processRolloverEntry( _
                ByVal pValue As String, _
                ByRef pEntryType As OrderTypes, _
                ByRef pEntryLimitPriceSpec As PriceSpecifier, _
                ByRef pEntryTriggerPriceSpec As PriceSpecifier, _
                ByRef pEntryTimeoutSecs As Long)
Const ProcName As String = "processRolloverEntry"
On Error GoTo Err

parseOrderSpec pValue, pEntryType, pEntryLimitPriceSpec, pEntryTriggerPriceSpec, pEntryTimeoutSecs

If Not IsEntryOrderType(pEntryType) Then
    pEntryType = OrderTypeNone
    Set pEntryLimitPriceSpec = New PriceSpecifier
    Set pEntryTriggerPriceSpec = New PriceSpecifier
    pEntryTimeoutSecs = 0
    AssertArgument True, "Invalid entry order type"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processRolloverQuantity( _
                ByVal pValue As String, _
                ByRef pQuantityMode As RolloverQuantityModes, _
                ByRef pQuantityParameter As BoxedDecimal, _
                ByRef pLotSize As Long)
Const ProcName As String = "processRolloverQuantity"
On Error GoTo Err

Const QuantityFormatLotSize As String = "(?:\(([1-9]\d*)\))?"
Const QuantityFormatNumber As String = "^([1-9]\d*)$"
Const QuantityFormatPreviousNumber As String = "^(<)$"
Const QuantityFormatMonetaryAmount As String = "^([1-9]\d*)\$" & QuantityFormatLotSize & "$"
Const QuantityFormatPreviousMonetaryAmount As String = "^(<\$)$"
Const QuantityFormatPercentOfAccount As String = "^([1-9]?\d*(?:\.\d+)?)%" & QuantityFormatLotSize & "$"
Const QuantityFormatPreviousPercentOfAccount As String = "^(<%)$"
Const QuantityFormatCurrentValue As String = "^(=\$)" & QuantityFormatLotSize & "$"
Const QuantityFormatCurrentProfit As String = "^(?:=([1-9]?\d*(?:\.\d+)?))%P" & QuantityFormatLotSize & "$"

Const QuantityFormat As String = QuantityFormatNumber & _
                            "|" & QuantityFormatPreviousNumber & _
                            "|" & QuantityFormatMonetaryAmount & _
                            "|" & QuantityFormatPreviousMonetaryAmount & _
                            "|" & QuantityFormatPercentOfAccount & _
                            "|" & QuantityFormatPreviousPercentOfAccount & _
                            "|" & QuantityFormatCurrentValue & _
                            "|" & QuantityFormatCurrentProfit

' QuantityFormat="^([1-9]\d*)$|^(<)$|^([1-9]\d*)\$(?:\(([1-9]\d*)\))?$|^(<\$)$|^([1-9]?\d*(?:\.\d+)?)%(?:\(([1-9]\d*)\))?$|^(<%)$|^(=\$)(?:\(([1-9]\d*)\))?$|^(?:=([1-9]?\d*(?:\.\d+)?))%P(?:\(([1-9]\d*)\))?$"

gRegExp.Pattern = QuantityFormat

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(pValue)

AssertArgument lMatches.Count = 1, "Invalid Quantity syntax"

Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lParameter As String

lParameter = lMatch.SubMatches(0)
If lParameter <> "" Then
    pQuantityMode = RolloverQuantityModeNumber
    Set pQuantityParameter = CreateBoxedDecimal(lParameter)
    Exit Sub
End If

If lMatch.SubMatches(1) <> "" Then
    pQuantityMode = RolloverQuantityModePreviousNumber
    Exit Sub
End If

lParameter = lMatch.SubMatches(2)
If lParameter <> "" Then
    pQuantityMode = RolloverQuantityModeMonetaryAmount
    Set pQuantityParameter = CreateBoxedDecimal(lParameter)
    pLotSize = CLng("0" & lMatch.SubMatches(3))
    Exit Sub
End If

If lMatch.SubMatches(4) <> "" Then
    pQuantityMode = RolloverQuantityModePreviousMonetaryAmount
    Exit Sub
End If

lParameter = lMatch.SubMatches(5)
If lParameter <> "" Then
    pQuantityMode = RolloverQuantityModePercentageOfAccount
    Set pQuantityParameter = CreateBoxedDecimal(lParameter)
    pLotSize = CLng("0" & lMatch.SubMatches(6))
    Exit Sub
End If

If lMatch.SubMatches(7) <> "" Then
    pQuantityMode = RolloverQuantityModePreviousPercentageOfAccount
    Exit Sub
End If

If lMatch.SubMatches(8) <> "" Then
    pQuantityMode = RolloverQuantityModeCurrentValue
    Set pQuantityParameter = CreateBoxedDecimal(lMatch.SubMatches(9))
    Exit Sub
End If

lParameter = lMatch.SubMatches(10)
If lParameter <> "" Then
    pQuantityMode = RolloverQuantityModeCurrentProfit
    Set pQuantityParameter = CreateBoxedDecimal(lParameter)
    pLotSize = CLng("0" & lMatch.SubMatches(11))
    Exit Sub
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processRolloverStrike( _
                ByVal pValue As String, _
                ByRef pStrikeMode As RolloverStrikeModes, _
                ByRef pStrikeValue As Double, _
                ByRef pStrikeOperator As OptionStrikeSelectionOperators, _
                ByRef pUnderlyingExchange As String)
Const ProcName As String = "processRolloverStrike"
On Error GoTo Err

Const StrikeFormatCurrentValue As String = "^(\$)$"
Const StrikeFormatCurrentProfit As String = "^([0-9]\d*)(?:%P)$"
Const StrikeFormatMonetaryAmount As String = "^([0-9]\d*)(?:\$)$"
Const StrikeFormatPreviousMonetaryAmount As String = "^(<\$)$"
Const StrikeFormatIncrement As String = "^([+-][1-9]\d?)$"
Const StrikeFormatDelta As String = "^(?:(?:(<|<=|>|>=|)(\-?[1-9]\d?)(D)(?:(?:;|,)([a-zA-Z0-9]+))?)?)?$"
Const StrikeFormatPreviousDelta As String = "<D"
Const StrikeFormatDeltaIncrement As String = "^(?:([+-][1-9]\d?)<D)$"


Const StrikeFormat As String = StrikeFormatCurrentValue & _
                            "|" & StrikeFormatCurrentProfit & _
                            "|" & StrikeFormatMonetaryAmount & _
                            "|" & StrikeFormatPreviousMonetaryAmount & _
                            "|" & StrikeFormatIncrement & _
                            "|" & StrikeFormatDelta & _
                            "|" & StrikeFormatPreviousDelta & _
                            "|" & StrikeFormatDeltaIncrement

gRegExp.Pattern = StrikeFormat

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(pValue)

AssertArgument lMatches.Count = 1, "Invalid " & StrikeSwitch & " syntax"

Dim lMatch As Match: Set lMatch = lMatches(0)

If lMatch.SubMatches(0) <> "" Then
    pStrikeMode = RolloverStrikeModeCurrentValue
    Exit Sub
End If

If lMatch.SubMatches(1) <> "" Then
    pStrikeMode = RolloverStrikeModeCurrentProfit
    pStrikeValue = CDbl(lMatch.SubMatches(1))
    Exit Sub
End If

If lMatch.SubMatches(2) <> "" Then
    pStrikeMode = RolloverStrikeModeMonetaryAmount
    pStrikeValue = CDbl(lMatch.SubMatches(2))
    Exit Sub
End If

If lMatch.SubMatches(3) <> "" Then
    pStrikeMode = RolloverStrikeModePreviousMonetaryAmount
    Exit Sub
End If

If lMatch.SubMatches(4) <> "" Then
    pStrikeMode = RolloverStrikeModeIncrement
    Exit Sub
End If

If lMatch.SubMatches(7) <> "" Then
    pStrikeMode = RolloverStrikeModeDelta
    pStrikeValue = CDbl(lMatch.SubMatches(6))
    pStrikeOperator = OptionStrikeSelectionOperatorFromString(lMatch.SubMatches(5))
    pUnderlyingExchange = lMatch.SubMatches(8)
    Exit Sub
End If

If lMatch.SubMatches(9) <> "" Then
    pStrikeMode = RolloverStrikeModePreviousDelta
    Exit Sub
End If

If lMatch.SubMatches(10) <> "" Then
    pStrikeMode = RolloverStrikeModeDeltaIncrement
    pStrikeValue = CDbl(lMatch.SubMatches(10))
    Exit Sub
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processRolloverTime(ByVal pValue As String) As Date
AssertArgument IsDate(pValue), "Time must be hh:mm[:ss]"
    
processRolloverTime = CDate(pValue)
AssertArgument CDbl(processRolloverTime) < 1#, "Time must be hh:mm[:ss]"


End Function




