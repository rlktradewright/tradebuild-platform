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
Public EntryPrice                                   As Double
Public EntryOffset                                  As Long
Public EntryTriggerPrice                            As Double
Public StopLossType                                 As BracketStopLossTypes
Public StopLossPrice                                As Double
Public StopLossOffset                               As Long
Public StopLossTriggerPrice                         As Double
Public TargetType                                   As BracketTargetTypes
Public TargetPrice                                  As Double
Public TargetOffset                                 As Long
Public TargetTriggerPrice                           As Double
Public EntryTIF                                     As OrderTIFs
Public StopTIF                                      As OrderTIFs
Public TargetTIF                                    As OrderTIFs
Public CancelPrice                                  As Double
Public CancelAfter                                  As Long

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
                gGenerateSwitch(PriceSwitch, EntryPrice)
    Case BracketEntryTypeLimit, _
            BracketEntryTypeLimitOnOpen, _
            BracketEntryTypeLimitOnClose
        s = s & gGenerateSwitch(PriceSwitch, EntryPrice)
    Case BracketEntryTypeBid, _
            BracketEntryTypeAsk, _
            BracketEntryTypeLast
        s = s & gGenerateSwitch(OffsetSwitch, EntryOffset)
End Select
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
                gGenerateSwitch(PriceSwitch, StopLossPrice)
    Case BracketStopLossTypeBid, _
            BracketStopLossTypeAsk, _
            BracketStopLossTypeAsk, _
            BracketStopLossTypeAuto
        s = s & gGenerateSwitch(OffsetSwitch, StopLossOffset)
End Select

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
    Case BracketTargetTypeMarketToLimit
    Case BracketTargetTypeLimit
        s = s & gGenerateSwitch(PriceSwitch, TargetPrice)
    Case BracketTargetTypeLimitIfTouched
        s = s & gGenerateSwitch(TriggerPriceSwitch, TargetTriggerPrice) & _
                gGenerateSwitch(PriceSwitch, TargetPrice)
    Case BracketTargetTypeMarketIfTouched
        s = s & gGenerateSwitch(TriggerPriceSwitch, TargetTriggerPrice)
    Case BracketTargetTypeBid, _
            BracketTargetTypeAsk, _
            BracketTargetTypeLast, _
            BracketTargetTypeAuto
        s = s & gGenerateSwitch(OffsetSwitch, TargetOffset)
End Select

generateTargetAttributes = s
End Function


