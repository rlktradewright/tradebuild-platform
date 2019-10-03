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

Public Action                                       As OrderActions
Public Quantity                                     As Long
Public EntryType                                    As BracketEntryTypes
Public EntryLimitPrice                              As Double
Public EntryLimitPriceOffset                        As Double
Public EntryLimitPriceOffsetType                    As PriceOffsetTypes
Public EntryTriggerPrice                            As Double
Public EntryTriggerPriceOffset                      As Double
Public EntryTriggerPriceOffsetType                  As PriceOffsetTypes
Public StopLossType                                 As BracketStopLossTypes
Public StopLossLimitPrice                           As Double
Public StopLossLimitPriceOffset                     As Long
Public StopLossLimitPriceOffsetType                 As PriceOffsetTypes
Public StopLossTriggerPrice                         As Double
Public StopLossTriggerPriceOffset                   As Double
Public StopLossTriggerPriceOffsetType               As PriceOffsetTypes
Public TargetType                                   As BracketTargetTypes
Public TargetLimitPrice                             As Double
Public TargetLimitPriceOffset                       As Long
Public TargetLimitPriceOffsetType                   As PriceOffsetTypes
Public TargetTriggerPrice                           As Double
Public TargetTriggerPriceOffset                     As Double
Public TargetTriggerPriceOffsetType                 As PriceOffsetTypes
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

'@================================================================================
' Class Event Handlers
'@================================================================================

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

Friend Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

Dim s As String
s = generateBracket

s = s & generateEntry
s = s & generateStopLoss
s = s & generateTarget

s = s & EndBracketCommand
ToString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function generateBracket() As String
Dim s As String
s = BracketCommand
s = s & " "

s = s & OrderActionToString(Action)
s = s & " "
s = s & CStr(Quantity)
s = s & " "

If CancelAfter <> 0 Then s = s & gGenerateSwitch(CancelAfterSwitch, CStr(CancelAfter))
If CancelPrice <> 0 Then s = s & gGenerateSwitch(CancelPriceSwitch, CStr(CancelPrice))

If GoodTillDate <> 0 Then s = s & gGenerateSwitch(GoodTillDateSwitch, Format(GoodTillDate, "yyyymmdd hh:mm:ss"))
If GoodAfterTime <> 0 Then s = s & gGenerateSwitch(GoodAfterTimeSwitch, Format(GoodAfterTime, "yyyymmdd hh:mm:ss"))

If TimezoneName <> "" Then s = s & gGenerateSwitch(TimezoneSwitch, TimezoneName)

generateBracket = s & vbCrLf
End Function

Private Function generateEntry() As String
If EntryType = BracketEntryTypeNone Then Exit Function

Dim s As String
s = EntryCommand & " "
s = s & BracketEntryTypeToShortString(EntryType) & " "
s = s & generateEntryAttributes
generateEntry = s & vbCrLf
End Function

Private Function generateEntryAttributes() As String
Dim s As String
Select Case EntryType
Case BracketEntryTypeMarket
Case BracketEntryTypeMarketOnOpen
Case BracketEntryTypeMarketOnClose
Case BracketEntryTypeMarketToLimit
Case BracketEntryTypeStop, _
        BracketEntryTypeMarketIfTouched
    s = s & gGenerateSwitch(TriggerPriceSwitch, EntryTriggerPrice)
Case BracketEntryTypeStopLimit, _
        BracketEntryTypeLimitIfTouched
    s = s & gGenerateSwitch(TriggerPriceSwitch, EntryTriggerPrice) & _
            gGenerateSwitch(PriceSwitch, EntryLimitPrice)
Case BracketEntryTypeLimit, _
        BracketEntryTypeLimitOnOpen, _
        BracketEntryTypeLimitOnClose
    s = s & gGenerateSwitch(PriceSwitch, EntryLimitPrice)
Case BracketEntryTypeBid, _
        BracketEntryTypeAsk, _
        BracketEntryTypeLast
    s = s & gGenerateSwitch(OffsetSwitch, gGenerateOffset(EntryLimitPriceOffset, EntryLimitPriceOffsetType))
End Select

If EntryTIF <> OrderTIFNone Then s = s & gGenerateSwitch(TIFSwitch, OrderTIFToString(EntryTIF))
generateEntryAttributes = s
End Function

Private Function generateStopLoss() As String
If StopLossType = BracketStopLossTypeNone Then Exit Function

Dim s As String
s = StopLossCommand & " "
s = s & BracketStopLossTypeToShortString(StopLossType) & " "
s = s & generateStopLossAttributes
generateStopLoss = s & vbCrLf
End Function

Private Function generateStopLossAttributes() As String
Dim s As String
Select Case StopLossType
Case BracketStopLossTypeStop
    s = s & gGenerateSwitch(TriggerPriceSwitch, StopLossTriggerPrice)
Case BracketStopLossTypeStopLimit
    s = s & gGenerateSwitch(TriggerPriceSwitch, StopLossTriggerPrice) & _
            gGenerateSwitch(PriceSwitch, StopLossLimitPrice)
Case BracketStopLossTypeBid, _
        BracketStopLossTypeAsk, _
        BracketStopLossTypeAsk, _
        BracketStopLossTypeAuto
    s = s & gGenerateSwitch(OffsetSwitch, gGenerateOffset(StopLossTriggerPriceOffset, StopLossTriggerPriceOffsetType))
End Select

If StopLossTIF <> OrderTIFNone Then s = s & gGenerateSwitch(TIFSwitch, OrderTIFToString(StopLossTIF))
generateStopLossAttributes = s
End Function

Private Function generateTarget() As String
If TargetType = BracketTargetTypeNone Then Exit Function

Dim s As String
s = TargetCommand & " "
s = s & BracketTargetTypeToShortString(TargetType) & " "
s = s & generateTargetAttributes
generateTarget = s & vbCrLf
End Function

Private Function generateTargetAttributes() As String
Dim s As String
Select Case TargetType
    Case BracketTargetTypeLimit
        s = s & gGenerateSwitch(PriceSwitch, TargetLimitPrice)
    Case BracketTargetTypeLimitIfTouched
        s = s & gGenerateSwitch(TriggerPriceSwitch, TargetTriggerPrice) & _
                gGenerateSwitch(PriceSwitch, TargetLimitPrice)
    Case BracketTargetTypeMarketIfTouched
        s = s & gGenerateSwitch(TriggerPriceSwitch, TargetTriggerPrice)
    Case BracketTargetTypeBid, _
            BracketTargetTypeAsk, _
            BracketTargetTypeLast, _
            BracketTargetTypeAuto
        s = s & gGenerateSwitch(OffsetSwitch, gGenerateOffset(TargetLimitPriceOffset, TargetLimitPriceOffsetType))
End Select

If TargetTIF <> OrderTIFNone Then s = s & gGenerateSwitch(TIFSwitch, OrderTIFToString(TargetTIF))
generateTargetAttributes = s
End Function


