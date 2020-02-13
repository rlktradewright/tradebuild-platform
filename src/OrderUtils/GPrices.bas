Attribute VB_Name = "GPrices"
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

Private Const ModuleName                            As String = "GPrices"

Public Const AskPriceDesignator                     As String = "ASK"
Public Const BidPriceDesignator                     As String = "BID"
Public Const BidAskPriceDesignator                  As String = "BIDASK"
Public Const LastPriceDesignator                    As String = "LAST"
Public Const EntryPriceDesignator                   As String = "ENTRY"
Public Const MidPriceDesignator                     As String = "MID"

Public Const TickOffsetDesignator                   As String = "T"
Public Const PercentOffsetDesignator                As String = "%"
Public Const BidAskSpreadPercentOffsetDesignator    As String = "S"

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

Public Function gNewPriceSpecifier( _
                Optional ByVal pPrice As Double = MaxDoubleValue, _
                Optional ByVal pPriceType As PriceValueTypes = PriceValueTypeNone, _
                Optional ByVal pOffset As Double = 0#, _
                Optional ByVal pOffsetType As PriceOffsetTypes = PriceOffsetTypeNone, _
                Optional ByVal pUseCloseoutSemantics As Boolean = False) As PriceSpecifier
Dim p As New PriceSpecifier
p.Initialise pPrice, pPriceType, pOffset, pOffsetType, pUseCloseoutSemantics
Set gNewPriceSpecifier = p
End Function

Public Function gParsePriceAndOffset( _
                ByRef pPriceSpec As PriceSpecifier, _
                ByVal pValue As String, _
                ByVal pSectype As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByVal pUseCloseoutSemantics As Boolean) As Boolean
Const ProcName As String = "gParsePrice"
On Error GoTo Err

If pValue = "" Then
    Set pPriceSpec = gNewPriceSpecifier(MaxDoubleValue, PriceValueTypeNone, 0, PriceOffsetTypeNone)
    gParsePriceAndOffset = True
    Exit Function
End If

gRegExp.Global = False
gRegExp.IgnoreCase = True

Dim p As String
p = "(?:^(?:(ASK|BID|BIDASK|LAST|ENTRY|MID|(?:[-+]?\d{1,6}(?:.\d{1,6})?))?)(?:\[(?:([-+]?\d{1,6}(?:.\d{1,6})?))([T%S]?)\])?$)"

gRegExp.Pattern = p

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(Trim$(pValue))

If lMatches.Count <> 1 Then
    gParsePriceAndOffset = False
    Exit Function
End If

Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lPrice As Double: lPrice = MaxDoubleValue
Dim lPriceType As PriceValueTypes
Dim lOffset As Double
Dim lOffsetType As PriceOffsetTypes

Dim lPricePart As String: lPricePart = UCase$(lMatch.SubMatches(0))
Select Case lPricePart
Case ""
    lPriceType = PriceValueTypeBidOrAsk
Case AskPriceDesignator
    lPriceType = PriceValueTypeAsk
Case BidPriceDesignator
    lPriceType = PriceValueTypeBid
Case BidAskPriceDesignator
    lPriceType = PriceValueTypeBidOrAsk
Case LastPriceDesignator
    lPriceType = PriceValueTypeLast
Case EntryPriceDesignator
    lPriceType = PriceValueTypeEntry
Case MidPriceDesignator
    lPriceType = PriceValueTypeMid
Case Else
    lPriceType = PriceValueTypeValue
    If Not ParsePrice(lPricePart, pSectype, pTickSize, lPrice) Then
        gParsePriceAndOffset = False
        Exit Function
    End If
End Select

Dim lOffsetPart As String: lOffsetPart = lMatch.SubMatches(1)
If lOffsetPart <> "" Then lOffset = CDbl(lOffsetPart)

Dim lOffsetDesignator As String: lOffsetDesignator = UCase$(lMatch.SubMatches(2))
Select Case lOffsetDesignator
Case ""
    lOffsetType = PriceOffsetTypeIncrement
Case TickOffsetDesignator
    lOffsetType = PriceOffsetTypeNumberOfTicks
Case PercentOffsetDesignator
    lOffsetType = PriceOffsetTypePercent
Case BidAskSpreadPercentOffsetDesignator
    lOffsetType = PriceOffsetTypeBidAskPercent
End Select

Set pPriceSpec = gNewPriceSpecifier(lPrice, lPriceType, lOffset, lOffsetType, pUseCloseoutSemantics)

gParsePriceAndOffset = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPriceOffsetToString( _
                ByVal pOffset As Double, _
                ByVal pOffsetType As PriceOffsetTypes)
Const ProcName As String = "gPriceOffsetToString"
On Error GoTo Err

gPriceOffsetToString = pOffset & gPriceOffsetTypeToString(pOffsetType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPriceOffsetTypeToString( _
                ByVal pOffsetType As PriceOffsetTypes)
Const ProcName As String = "gPriceOffsetTypeToString"
On Error GoTo Err

Select Case pOffsetType
Case PriceOffsetTypeNone
    gPriceOffsetTypeToString = ""
Case PriceOffsetTypeIncrement
    gPriceOffsetTypeToString = ""
Case PriceOffsetTypeNumberOfTicks
    gPriceOffsetTypeToString = "T"
Case PriceOffsetTypeBidAskPercent
    gPriceOffsetTypeToString = "S"
Case PriceOffsetTypePercent
    gPriceOffsetTypeToString = "%"
Case Else
    AssertArgument False, "Value is not a valid Price Offset Type"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPriceOrSpecifierToString( _
                ByVal pPrice As Double, _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pContract As IContract) As String
Const ProcName As String = "gPriceOrSpecifierToString"
On Error GoTo Err

If pPrice = MaxDoubleValue Then
    gPriceOrSpecifierToString = gPriceSpecifierToString(pPriceSpec, pContract)
Else
    gPriceOrSpecifierToString = FormatPrice(pPrice, pContract.Specifier.SecType, pContract.TickSize)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPriceSpecifierToString( _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pContract As IContract) As String
Const ProcName As String = "gPriceSpecifierToString"
On Error GoTo Err

gPriceSpecifierToString = gTypedPriceToString( _
                                pPriceSpec.Price, _
                                pPriceSpec.PriceType, _
                                pContract) & _
                            "[" & _
                            gPriceOffsetToString(pPriceSpec.Offset, pPriceSpec.OffsetType) & _
                            "]"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gRoundToTickBoundary( _
                ByVal pPrice As Double, _
                ByVal pTickSize As Double, _
                ByVal pMode As TickRoundingModes) As Double

Dim lTicksPerUnit As Long: lTicksPerUnit = Round(1# / pTickSize)
Dim lErrorIncrement As Double: lErrorIncrement = lTicksPerUnit / 10000#

Select Case pMode
Case TickRoundingModeNearest
    gRoundToTickBoundary = Round(pPrice * lTicksPerUnit + lErrorIncrement) / lTicksPerUnit
Case TickRoundingModeDown
    gRoundToTickBoundary = Int(pPrice * lTicksPerUnit + lErrorIncrement) / lTicksPerUnit
Case TickRoundingModeUp
    gRoundToTickBoundary = -Int(-pPrice * lTicksPerUnit + lErrorIncrement) / lTicksPerUnit
End Select
End Function

Public Function gTypedPriceToString( _
                ByVal pPrice As Double, _
                ByVal pPriceType As PriceValueTypes, _
                ByVal pContract As IContract) As String
Const ProcName As String = "gTypedPriceToString"
On Error GoTo Err

Dim s As String

Select Case pPriceType
Case PriceValueTypeNone
Case PriceValueTypeValue
    If pPrice <> MaxDouble Then
        s = FormatPrice(pPrice, pContract.Specifier.SecType, pContract.TickSize)
    End If
Case PriceValueTypeAsk
    s = "ASK"
Case PriceValueTypeBid
    s = "BID"
Case PriceValueTypeLast
    s = "LAST"
Case PriceValueTypeEntry
    s = "ENTRY"
Case PriceValueTypeMid
    s = "MID"
Case PriceValueTypeBidOrAsk
    s = "BIDASK"
Case Else
    AssertArgument False, "Invalid price value type"
End Select

gTypedPriceToString = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




