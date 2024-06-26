VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderSpecification"
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

Implements IStringable

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

Private Const ModuleName                            As String = "BracketOrderSpecification"

'@================================================================================
' Member variables
'@================================================================================

Public ID                                           As String
Public Modify                                       As Boolean
Public Contract                                     As IContract
Public Action                                       As OrderActions
Public Quantity                                     As BoxedDecimal
Public AutosizeType                                 As BracketAutosizeTypes
Public AutosizeValue                                As Double
Public LotSize                                      As Long
Public EntryType                                    As OrderTypes
Public EntryLimitPriceSpec                          As PriceSpecifier
Public EntryTriggerPriceSpec                        As PriceSpecifier
Public EntryIgnoreRTH                               As Boolean
Public StopLossType                                 As OrderTypes
Public StopLossLimitPriceSpec                       As PriceSpecifier
Public StopLossTriggerPriceSpec                     As PriceSpecifier
Public StopLossIgnoreRTH                            As Boolean
Public TargetType                                   As OrderTypes
Public TargetLimitPriceSpec                         As PriceSpecifier
Public TargetTriggerPriceSpec                       As PriceSpecifier
Public TargetIgnoreRTH                              As Boolean
Public EntryTIF                                     As OrderTIFs
Public StopLossTIF                                  As OrderTIFs
Public TargetTIF                                    As OrderTIFs
Public CancelPrice                                  As Double
Public CancelAfter                                  As Long
Public GoodAfterTime                                As Date
Public GoodTillDate                                 As Date
Public TimezoneName                                 As String
Public Description                                  As String
Public EntryReason                                  As String
Public StopLossReason                               As String
Public TargetReason                                 As String
Public RolloverSpec                                 As RolloverSpecification

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set EntryLimitPriceSpec = New PriceSpecifier
Set EntryTriggerPriceSpec = New PriceSpecifier
Set StopLossLimitPriceSpec = New PriceSpecifier
Set StopLossTriggerPriceSpec = New PriceSpecifier
Set TargetLimitPriceSpec = New PriceSpecifier
Set TargetTriggerPriceSpec = New PriceSpecifier
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
Const ProcName As String = "IStringable_ToString"
On Error GoTo Err

IStringable_ToString = ToString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function ToSummaryString()
Const ProcName As String = "ToSummaryString"
On Error GoTo Err

Dim s As String
s = gGetContractName(Contract) & " " & _
        IIf(Action = OrderActionBuy, "B ", "S ") & _
        Quantity.DecimalValue & " " & _
        OrderTypeToShortString(EntryType) & " " & _
        EntryLimitPriceSpec.PriceString
If EntryTriggerPriceSpec.IsValid Then s = s & ";" & EntryTriggerPriceSpec.PriceString

If StopLossType <> OrderTypeNone Then
    s = s & " (SL: " & OrderTypeToShortString(StopLossType) & " "
    If StopLossLimitPriceSpec.IsValid Then _
        s = s & StopLossLimitPriceSpec.PriceString
    If StopLossTriggerPriceSpec.IsValid Then _
        s = s & ";" & StopLossTriggerPriceSpec.PriceString
    s = s & ")"
End If

If TargetType <> OrderTypeNone Then
    s = s & " (T: " & OrderTypeToShortString(TargetType) & " "
    If TargetLimitPriceSpec.IsValid Then _
        s = s & TargetLimitPriceSpec.PriceString
    If TargetTriggerPriceSpec.IsValid Then _
        s = s & ";" & TargetTriggerPriceSpec.PriceString
    s = s & ")"
End If

ToSummaryString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

If isSingleLine Then
    ToString = generateSingleLineOrder
Else
    ToString = generateBracketOrder
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatPriceSpec(ByVal pPriceSpec As PriceSpecifier) As String
formatPriceSpec = PriceSpecifierToString(pPriceSpec, Contract)
End Function

Private Function generateActionAndQuantity()
Dim s As String
s = s & OrderActionToString(Action)
s = s & " "
If AutosizeType = BracketAutosizeNone Then
    s = s & CStr(Quantity)
ElseIf AutosizeType = BracketAutosizeMonetaryAmount Then
    s = s & CStr(AutosizeValue) & MaxOrderCostSuffix
Else
    s = s & CStr(AutosizeValue) & AccountPercentSuffix
End If
If LotSize <> 0 Then s = s & "(" & LotSize & ")"
generateActionAndQuantity = s
End Function

Private Function generateBracket() As String
Dim s As String
s = gCommands.BracketCommand.Name
s = s & " "

s = s & generateActionAndQuantity
s = s & " "

If CancelAfter <> 0 Then s = s & gGenerateSwitch(CancelAfterSwitch, CStr(CancelAfter))
If CancelPrice <> 0 Then s = s & gGenerateSwitch(CancelPriceSwitch, CStr(CancelPrice))

If GoodTillDate <> 0 Then s = s & gGenerateSwitch(GoodTillDateSwitch, Format(GoodTillDate, "yyyymmdd hh:mm:ss"))
If GoodAfterTime <> 0 Then s = s & gGenerateSwitch(GoodAfterTimeSwitch, Format(GoodAfterTime, "yyyymmdd hh:mm:ss"))

If TimezoneName <> "" Then s = s & gGenerateSwitch(TimezoneSwitch, TimezoneName)

generateBracket = s & vbCrLf
End Function

Private Function generateBracketOrder() As String
Const ProcName As String = "generateBracketOrder"
On Error GoTo Err

Dim s As String
s = generateBracket
s = s & generateEntry
s = s & generateStopLoss
s = s & generateTarget
s = s & generateRollover
s = s & gCommands.EndBracketCommand.Name
generateBracketOrder = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateEntry() As String
If EntryType = OrderTypeNone Then Exit Function

Dim s As String
s = gCommands.EntryCommand.Name & " "
s = s & OrderTypeToShortString(EntryType) & " "
s = s & generateEntryAttributes
generateEntry = s & vbCrLf
End Function

Private Function generateEntryAttributes() As String
Dim s As String
Select Case EntryType
Case OrderTypeMarket
Case OrderTypeMidprice
Case OrderTypeMarketOnOpen
Case OrderTypeMarketOnClose
Case OrderTypeMarketToLimit
Case OrderTypeStop, _
        OrderTypeMarketIfTouched
    s = gGenerateSwitch(TriggerPriceSwitch, formatPriceSpec(EntryTriggerPriceSpec))
Case OrderTypeStopLimit, _
        OrderTypeLimitIfTouched
    s = gGenerateSwitch(TriggerPriceSwitch, formatPriceSpec(EntryTriggerPriceSpec)) & _
            gGenerateSwitch(PriceSwitch, formatPriceSpec(EntryLimitPriceSpec))
Case OrderTypeLimit, _
        OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose
    s = gGenerateSwitch(PriceSwitch, formatPriceSpec(EntryLimitPriceSpec))
Case OrderTypeTrail
    s = gGenerateSwitch(TriggerPriceSwitch, formatPriceSpec(EntryTriggerPriceSpec))
Case OrderTypeTrailLimit
    s = gGenerateSwitch(TriggerPriceSwitch, formatPriceSpec(EntryTriggerPriceSpec)) & _
            gGenerateSwitch(PriceSwitch, formatPriceSpec(EntryLimitPriceSpec))
End Select

If EntryTIF <> OrderTIFNone Then s = s & gGenerateSwitch(TIFSwitch, OrderTIFToString(EntryTIF))
If EntryIgnoreRTH Then s = s & gGenerateSwitch(IgnoreRTHSwitch, "")

generateEntryAttributes = s
End Function

Private Function generateRollover() As String
If RolloverSpec Is Nothing Then Exit Function

Dim s As String
s = gCommands.RolloverCommand.Name & " "

s = s & gGenerateSwitch(DaysSwitch, RolloverSpec.Days)
s = s & gGenerateSwitch(TimeSwitch, FormatTimestamp(RolloverSpec.Time, TimestampTimeOnlyISO8601 + TimestampNoMillisecs))
s = s & gGenerateSwitch(CloseSwitch, generateRolloverClose)
s = s & gGenerateSwitch(EntrySwitch, generateRolloverEntry)

generateRollover = s & vbCrLf
End Function

Private Function generateRolloverClose() As String
Dim s As String

s = OrderTypeToShortString(RolloverSpec.CloseOrderType) & ";"

Select Case RolloverSpec.CloseOrderType
Case OrderTypeMarket, _
        OrderTypeMarketOnOpen, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketToLimit
Case OrderTypeStop, _
        OrderTypeMarketIfTouched
    s = s & formatPriceSpec(RolloverSpec.CloseTriggerPriceSpec) & ";"
Case OrderTypeStopLimit, _
        OrderTypeLimitIfTouched
    s = s & formatPriceSpec(RolloverSpec.CloseLimitPriceSpec) & ";"
    s = s & formatPriceSpec(RolloverSpec.CloseTriggerPriceSpec) & ";"
Case OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose
Case OrderTypeLimit
    s = s & formatPriceSpec(RolloverSpec.CloseLimitPriceSpec) & ";"
Case OrderTypeTrail
    s = s & formatPriceSpec(RolloverSpec.CloseTriggerPriceSpec) & ";"
Case OrderTypeTrailLimit
    s = s & formatPriceSpec(RolloverSpec.CloseLimitPriceSpec) & ";"
    s = s & formatPriceSpec(RolloverSpec.CloseTriggerPriceSpec) & ";"
End Select

If RolloverSpec.CloseTimeoutSecs <> 0 Then s = s & RolloverSpec.CloseTimeoutSecs

Do While Right$(s, 1) = ";"
    s = Left$(s, Len(s) - 1)
Loop
generateRolloverClose = s
End Function

Private Function generateRolloverEntry() As String
Dim s As String

s = OrderTypeToShortString(RolloverSpec.EntryOrderType) & ";"

Select Case RolloverSpec.CloseOrderType
Case OrderTypeMarket, _
        OrderTypeMarketOnOpen, _
        OrderTypeMarketOnClose, _
        OrderTypeMarketToLimit
Case OrderTypeStop, _
        OrderTypeMarketIfTouched
    s = s & formatPriceSpec(RolloverSpec.EntryTriggerPriceSpec) & ";"
Case OrderTypeStopLimit, _
        OrderTypeLimitIfTouched
    s = s & formatPriceSpec(RolloverSpec.EntryLimitPriceSpec) & ";"
    s = s & formatPriceSpec(RolloverSpec.EntryTriggerPriceSpec) & ";"
Case OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose
Case OrderTypeLimit
    s = s & formatPriceSpec(RolloverSpec.EntryLimitPriceSpec) & ";"
Case OrderTypeTrail
    s = s & formatPriceSpec(RolloverSpec.EntryTriggerPriceSpec) & ";"
Case OrderTypeTrailLimit
    s = s & formatPriceSpec(RolloverSpec.EntryLimitPriceSpec) & ";"
    s = s & formatPriceSpec(RolloverSpec.EntryTriggerPriceSpec) & ";"
End Select

If RolloverSpec.EntryTimeoutSecs <> 0 Then s = s & RolloverSpec.EntryTimeoutSecs

Do While Right$(s, 1) = ";"
    s = Left$(s, Len(s) - 1)
Loop
generateRolloverEntry = s
End Function

Private Function generateSingleLineOrder() As String
Const ProcName As String = "generateSingleLineOrder"
On Error GoTo Err

Dim s As String
s = generateActionAndQuantity
s = s & " "
s = s & OrderTypeToShortString(EntryType) & " "
s = s & generateEntryAttributes

generateSingleLineOrder = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateSingleLineParameters() As String
Dim s As String
Select Case EntryType
Case OrderTypeMarket
Case OrderTypeMidprice
Case OrderTypeMarketOnOpen
Case OrderTypeMarketOnClose
Case OrderTypeMarketToLimit
Case OrderTypeStop, _
        OrderTypeMarketIfTouched
    s = formatPriceSpec(EntryTriggerPriceSpec)
Case OrderTypeStopLimit, _
        OrderTypeLimitIfTouched
    s = formatPriceSpec(EntryTriggerPriceSpec) & _
            formatPriceSpec(EntryLimitPriceSpec)
Case OrderTypeLimit, _
        OrderTypeLimitOnOpen, _
        OrderTypeLimitOnClose
    s = formatPriceSpec(EntryLimitPriceSpec)
Case OrderTypeTrail
    s = formatPriceSpec(EntryTriggerPriceSpec)
Case OrderTypeTrailLimit
    s = formatPriceSpec(EntryTriggerPriceSpec) & _
            formatPriceSpec(EntryLimitPriceSpec)
End Select

If EntryTIF <> OrderTIFNone Then s = s & gGenerateSwitch(TIFSwitch, OrderTIFToString(EntryTIF))
If EntryIgnoreRTH Then s = s & gGenerateSwitch(IgnoreRTHSwitch, "")

generateSingleLineParameters = s
End Function

Private Function generateStopLoss() As String
If StopLossType = OrderTypeNone Then Exit Function

Dim s As String
s = gCommands.StopLossCommand.Name & " "
s = s & OrderTypeToShortString(StopLossType) & " "
s = s & generateStopLossAttributes
generateStopLoss = s & vbCrLf
End Function

Private Function generateStopLossAttributes() As String
Dim s As String
Select Case StopLossType
Case OrderTypeStop, _
        OrderTypeTrail, _
        OrderTypeTrailLimit
    s = s & gGenerateSwitch(TriggerPriceSwitch, formatPriceSpec(StopLossTriggerPriceSpec))
Case OrderTypeStopLimit
    s = s & gGenerateSwitch(TriggerPriceSwitch, formatPriceSpec(StopLossTriggerPriceSpec)) & _
            gGenerateSwitch(PriceSwitch, formatPriceSpec(StopLossLimitPriceSpec))
End Select

If StopLossTIF <> OrderTIFNone Then s = s & gGenerateSwitch(TIFSwitch, OrderTIFToString(StopLossTIF))
If StopLossIgnoreRTH Then s = s & gGenerateSwitch(IgnoreRTHSwitch, "")
generateStopLossAttributes = s
End Function

Private Function generateTarget() As String
If TargetType = OrderTypeNone Then Exit Function

Dim s As String
s = gCommands.TargetCommand.Name & " "
s = s & OrderTypeToShortString(TargetType) & " "
s = s & generateTargetAttributes
generateTarget = s & vbCrLf
End Function

Private Function generateTargetAttributes() As String
Dim s As String
Select Case TargetType
    Case OrderTypeLimit
        s = s & gGenerateSwitch(PriceSwitch, formatPriceSpec(TargetLimitPriceSpec))
    Case OrderTypeLimitIfTouched
        s = s & gGenerateSwitch(TriggerPriceSwitch, formatPriceSpec(TargetTriggerPriceSpec)) & _
                gGenerateSwitch(PriceSwitch, formatPriceSpec(TargetLimitPriceSpec))
    Case OrderTypeMarketIfTouched
        s = s & gGenerateSwitch(TriggerPriceSwitch, formatPriceSpec(TargetTriggerPriceSpec))
End Select

If TargetTIF <> OrderTIFNone Then s = s & gGenerateSwitch(TIFSwitch, OrderTIFToString(TargetTIF))
If TargetIgnoreRTH Then s = s & gGenerateSwitch(IgnoreRTHSwitch, "")
generateTargetAttributes = s
End Function

Private Function isSingleLine() As Boolean
isSingleLine = (StopLossType = OrderTypeNone And TargetType = OrderTypeNone)
End Function


