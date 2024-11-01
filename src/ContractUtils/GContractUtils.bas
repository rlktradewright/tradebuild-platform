Attribute VB_Name = "GContractUtils"
Option Explicit

''
' Description here
'
' @remarks
' @see
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

Private Const ModuleName                    As String = "GContractUtils"

Public Const ConfigSectionContractSpecifier             As String = "Specifier"

Public Const ConfigSettingContractSpecCurrency          As String = "&Currency"
Public Const ConfigSettingContractSpecExpiry            As String = "&Expiry"
Public Const ConfigSettingContractSpecExchange          As String = "&Exchange"
Public Const ConfigSettingContractSpecLocalSymbol       As String = "&LocalSymbol"
Public Const ConfigSettingContractSpecMultiplier        As String = "&Multiplier"
Public Const ConfigSettingContractSpecRight             As String = "&Right"
Public Const ConfigSettingContractSpecSecType           As String = "&SecType"
Public Const ConfigSettingContractSpecStrikePrice       As String = "&StrikePrice"
Public Const ConfigSettingContractSpecSymbol            As String = "&Symbol"
Public Const ConfigSettingContractSpecTradingClass      As String = "&TradingClass"

Public Const ConfigSettingDaysBeforeExpiryToSwitch      As String = "&DaysBeforeExpiryToSwitch"
Public Const ConfigSettingDescription                   As String = "&Description"
Public Const ConfigSettingExpiryDate                    As String = "&ExpiryDate"
Public Const ConfigSettingFullSessionEndTime            As String = "&FullSessionEndTime"
Public Const ConfigSettingFullSessionStartTime          As String = "&FullSessionStartTime"
Public Const ConfigSettingMultiplier                    As String = "&Multiplier"
Public Const ConfigSettingSessionEndTime                As String = "&SessionEndTime"
Public Const ConfigSettingSessionStartTime              As String = "&SessionStartTime"
Public Const ConfigSettingTickSize                      As String = "&TickSize"
Public Const ConfigSettingTimezoneName                  As String = "&Timezone"

Public Const DefaultDaysBeforeExpiryToSwitch            As Long = 0
Public Const DefaultExpiry                              As String = "1899-12-30"
Public Const DefaultMultiplier                          As Double = 0#
Public Const DefaultTickSize                            As Double = 0.01
Public Const DefaultTimezoneName                        As String = "Eastern Standard Time"

Public Const MaxContractExpiryOffset                    As Long = 10
Public Const MaxContractDaysBeforeExpiryToSwitch        As Long = 20

'@================================================================================
' Member variables
'@================================================================================

Private mExchangeCodesInitialised                       As Boolean
Private mExchangeCodes()                                As String
Private mMaxExchangeCodesIndex                          As Long

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

Public Property Let ExactSixtyFourthIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gExactSixtyFourthIndicators = ar
End Property
                
Public Property Get ExactSixtyFourthIndicators() As String()
ExactSixtyFourthIndicators = gExactSixtyFourthIndicators
End Property
                
Public Property Let ExactThirtySecondIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gExactThirtySecondIndicators = ar
End Property
                
Public Property Get ExactThirtySecondIndicators() As String()
ExactThirtySecondIndicators = gExactThirtySecondIndicators
End Property
                
Public Property Let HalfSixtyFourthIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gHalfSixtyFourthIndicators = ar
End Property
                
Public Property Get HalfSixtyFourthIndicators() As String()
HalfSixtyFourthIndicators = gHalfSixtyFourthIndicators
End Property
                
Public Property Let HalfThirtySecondIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gHalfThirtySecondIndicators = ar
End Property
                
Public Property Get HalfThirtySecondIndicators() As String()
HalfThirtySecondIndicators = gHalfThirtySecondIndicators
End Property

Public Property Let QuarterThirtySecondIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gQuarterThirtySecondIndicators = ar
End Property
                
Public Property Get QuarterThirtySecondIndicators() As String()
QuarterThirtySecondIndicators = gQuarterThirtySecondIndicators
End Property
                
Public Property Let SixtyFourthsAndFractionsSeparators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gSixtyFourthsAndFractionsSeparators = ar
End Property
                
Public Property Get SixtyFourthsAndFractionsSeparators() As String()
SixtyFourthsAndFractionsSeparators = gSixtyFourthsAndFractionsSeparators
End Property
                
Public Property Let SixtyFourthsAndFractionsTerminators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gSixtyFourthsAndFractionsTerminators = ar
End Property
                
Public Property Get SixtyFourthsAndFractionsTerminators() As String()
SixtyFourthsAndFractionsTerminators = gSixtyFourthsAndFractionsTerminators
End Property
                
Public Property Let SixtyFourthsSeparators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gSixtyFourthsSeparators = ar
End Property
                
Public Property Get SixtyFourthsSeparators() As String()
SixtyFourthsSeparators = gSixtyFourthsSeparators
End Property
                
Public Property Let SixtyFourthsTerminators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gSixtyFourthsTerminators = ar
End Property
                
Public Property Get SixtyFourthsTerminators() As String()
SixtyFourthsTerminators = gSixtyFourthsTerminators
End Property
                
Public Property Let ThirtySecondsAndFractionsSeparators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThirtySecondsAndFractionsSeparators = ar
End Property
                
Public Property Get ThirtySecondsAndFractionsSeparators() As String()
ThirtySecondsAndFractionsSeparators = gThirtySecondsAndFractionsSeparators
End Property
                
Public Property Let ThirtySecondsAndFractionsTerminators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThirtySecondsAndFractionsTerminators = ar
End Property
                
Public Property Get ThirtySecondsAndFractionsTerminators() As String()
ThirtySecondsAndFractionsTerminators = gThirtySecondsAndFractionsTerminators
End Property
                
Public Property Let ThirtySecondsSeparators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThirtySecondsSeparators = ar
End Property
                
Public Property Get ThirtySecondsSeparators() As String()
ThirtySecondsSeparators = gThirtySecondsSeparators
End Property
                
Public Property Let ThirtySecondsTerminators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThirtySecondsTerminators = ar
End Property
                
Public Property Get ThirtySecondsTerminators() As String()
ThirtySecondsTerminators = gThirtySecondsTerminators
End Property
                
Public Property Let ThreeQuarterThirtySecondIndicators( _
                ByRef Value() As String)
Dim ar() As String
ar = Value
gThreeQuarterThirtySecondIndicators = ar
End Property

Public Property Get ThreeQuarterThirtySecondIndicators() As String()
ThreeQuarterThirtySecondIndicators = gThreeQuarterThirtySecondIndicators
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function ContractSpecsCompare( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                ByVal pAscending As Boolean) As Long
Const ProcName As String = "ContractSpecsCompare"
On Error GoTo Err

Dim i As Long
For i = 0 To UBound(pSortkeys)
    Select Case pSortkeys(i)
    Case ContractSortKeyNone
        Exit Function
    Case ContractSortKeyLocalSymbol
        ContractSpecsCompare = StrComp(pContractSpec1.LocalSymbol, pContractSpec2.LocalSymbol, vbTextCompare)
    Case ContractSortKeySymbol
        ContractSpecsCompare = StrComp(pContractSpec1.Symbol, pContractSpec2.Symbol, vbTextCompare)
    Case ContractSortKeySecType
        ContractSpecsCompare = StrComp(GContractUtils.SecTypeToShortString(pContractSpec1.SecType), GContractUtils.SecTypeToShortString(pContractSpec2.SecType), vbTextCompare)
    Case ContractSortKeyExchange
        ContractSpecsCompare = StrComp(pContractSpec1.Exchange, pContractSpec2.Exchange, vbTextCompare)
    Case ContractSortKeyExpiry
        ContractSpecsCompare = StrComp(pContractSpec1.Expiry, pContractSpec2.Expiry, vbTextCompare)
    Case ContractSortKeyMultiplier
        ContractSpecsCompare = StrComp(pContractSpec1.Multiplier, pContractSpec2.Multiplier, vbBinaryCompare)
    Case ContractSortKeyCurrency
        ContractSpecsCompare = StrComp(pContractSpec1.CurrencyCode, pContractSpec2.CurrencyCode, vbTextCompare)
    Case ContractSortKeyRight
        ContractSpecsCompare = StrComp(GContractUtils.OptionRightToString(pContractSpec1.Right), GContractUtils.OptionRightToString(pContractSpec2.Right), vbTextCompare)
    Case ContractSortKeyStrike
        ContractSpecsCompare = Sgn(pContractSpec1.Strike - pContractSpec2.Strike)
    End Select
    If ContractSpecsCompare <> 0 Then
        If Not pAscending Then ContractSpecsCompare = -ContractSpecsCompare
        Exit Function
    End If
Next

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractSpecsCompatible( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
Const ProcName As String = "ContractSpecsCompatible"
On Error GoTo Err

If pContractSpec1.LocalSymbol <> "" And UCase$(pContractSpec1.LocalSymbol) <> UCase$(pContractSpec2.LocalSymbol) Then Exit Function
If pContractSpec1.Symbol <> "" And UCase$(pContractSpec1.Symbol) <> UCase$(pContractSpec2.Symbol) Then Exit Function
If pContractSpec1.SecType <> SecTypeNone And pContractSpec1.SecType <> pContractSpec2.SecType Then Exit Function
If pContractSpec1.Exchange <> "" And UCase$(pContractSpec1.Exchange) <> UCase$(pContractSpec2.Exchange) Then Exit Function
If pContractSpec1.Expiry <> "" And pContractSpec1.Expiry <> Left$(pContractSpec2.Expiry, Len(pContractSpec1.Expiry)) Then Exit Function
If pContractSpec1.Multiplier <> pContractSpec2.Multiplier Then Exit Function
If pContractSpec1.CurrencyCode <> "" And UCase$(pContractSpec1.CurrencyCode) <> UCase$(pContractSpec2.CurrencyCode) Then Exit Function
If pContractSpec1.Right <> OptNone And pContractSpec1.Right <> pContractSpec2.Right Then Exit Function
If pContractSpec1.Strike <> 0# And pContractSpec1.Strike <> pContractSpec2.Strike Then Exit Function

ContractSpecsCompatible = True

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractSpecsEqual( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
Const ProcName As String = "ContractSpecsEqual"
On Error GoTo Err

If pContractSpec1 Is Nothing Then Exit Function
If pContractSpec2 Is Nothing Then Exit Function
If pContractSpec1 Is pContractSpec2 Then
    ContractSpecsEqual = True
Else
    ContractSpecsEqual = (GContractUtils.GetContractSpecKey(pContractSpec1) = GContractUtils.GetContractSpecKey(pContractSpec2))
    If pContractSpec1.CurrencyCode <> pContractSpec2.CurrencyCode Then Exit Function
    If pContractSpec1.Exchange <> pContractSpec2.Exchange Then Exit Function
    If pContractSpec1.Expiry <> pContractSpec2.Expiry Then Exit Function
    If pContractSpec1.LocalSymbol <> pContractSpec2.LocalSymbol Then Exit Function
    If pContractSpec1.Multiplier <> pContractSpec2.Multiplier Then Exit Function
    If pContractSpec1.Right <> pContractSpec2.Right Then Exit Function
    If pContractSpec1.SecType <> pContractSpec2.SecType Then Exit Function
    If pContractSpec1.Strike <> pContractSpec2.Strike Then Exit Function
    If pContractSpec1.Symbol <> pContractSpec2.Symbol Then Exit Function
    If pContractSpec1.TradingClass <> pContractSpec2.TradingClass Then Exit Function
    ContractSpecsEqual = True
End If

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractToString(ByVal pContract As IContract) As String
Const ProcName As String = "ContractToString"
On Error GoTo Err

ContractToString = "Specifier=(" & pContract.Specifier.ToString & "); " & _
            "Description=" & pContract.Description & "; " & _
            "Expiry date=" & pContract.ExpiryDate & "; " & _
            "Tick size=" & pContract.TickSize & "; " & _
            "Session start=" & FormatDateTime(pContract.SessionStartTime, vbShortTime) & "; " & _
            "Session end=" & FormatDateTime(pContract.SessionEndTime, vbShortTime) & "; " & _
            "Full session start=" & FormatDateTime(pContract.FullSessionStartTime, vbShortTime) & "; " & _
            "Full session end=" & FormatDateTime(pContract.FullSessionEndTime, vbShortTime) & "; " & _
            "TimezoneName=" & pContract.TimezoneName

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ContractToXML(ByVal pContract As IContract) As String
Const ProcName As String = "ContractToXML"
On Error GoTo Err

Dim XMLdoc As DOMDocument60: Set XMLdoc = New DOMDocument60
Dim lContractElement As IXMLDOMElement: Set lContractElement = XMLdoc.createElement("contract")
Set XMLdoc.documentElement = lContractElement
lContractElement.setAttribute "xmlns", "urn:tradewright.com:tradebuild"
lContractElement.setAttribute "minimumtick", pContract.TickSize
lContractElement.setAttribute "sessionstarttime", Format(pContract.SessionStartTime, "hh:mm:ss")
lContractElement.setAttribute "sessionendtime", Format(pContract.SessionEndTime, "hh:mm:ss")
lContractElement.setAttribute "fullsessionstarttime", Format(pContract.FullSessionStartTime, "hh:mm:ss")
lContractElement.setAttribute "fullsessionendtime", Format(pContract.FullSessionEndTime, "hh:mm:ss")
lContractElement.setAttribute "Description", pContract.Description
lContractElement.setAttribute "numberofdecimals", pContract.NumberOfDecimals
lContractElement.setAttribute "timezonename", pContract.TimezoneName

Dim Specifier As IXMLDOMElement: Set Specifier = XMLdoc.createElement("specifier")
lContractElement.appendChild Specifier
Specifier.setAttribute "symbol", pContract.Specifier.Symbol
Specifier.setAttribute "tradingclass", pContract.Specifier.TradingClass
Specifier.setAttribute "sectype", pContract.Specifier.SecType
Specifier.setAttribute "expiry", pContract.Specifier.Expiry
Specifier.setAttribute "exchange", pContract.Specifier.Exchange
Specifier.setAttribute "currencycode", pContract.Specifier.CurrencyCode
Specifier.setAttribute "localsymbol", pContract.Specifier.LocalSymbol
Specifier.setAttribute "multiplier", pContract.Specifier.Multiplier
Specifier.setAttribute "right", pContract.Specifier.Right
Specifier.setAttribute "strike", pContract.Specifier.Strike

Dim ComboLegs As IXMLDOMElement: Set ComboLegs = XMLdoc.createElement("combolegs")
Specifier.appendChild ComboLegs
If Not pContract.Specifier.ComboLegs Is Nothing Then
    Dim comboLegObj As ComboLeg
    For Each comboLegObj In pContract.Specifier.ComboLegs
        Dim ComboLeg As IXMLDOMElement: Set ComboLeg = XMLdoc.createElement("comboleg")
        ComboLegs.appendChild ComboLeg
        
        Set Specifier = XMLdoc.createElement("specifier")
        ComboLeg.appendChild Specifier
        ComboLeg.setAttribute "isBuyLeg", comboLegObj.IsBuyLeg
        ComboLeg.setAttribute "Ratio", comboLegObj.Ratio
        Specifier.setAttribute "symbol", comboLegObj.ContractSpec.Symbol
        Specifier.setAttribute "tradingclass", comboLegObj.ContractSpec.TradingClass
        Specifier.setAttribute "sectype", comboLegObj.ContractSpec.SecType
        Specifier.setAttribute "expiry", comboLegObj.ContractSpec.Expiry
        Specifier.setAttribute "exchange", comboLegObj.ContractSpec.Exchange
        Specifier.setAttribute "currencycode", comboLegObj.ContractSpec.CurrencyCode
        Specifier.setAttribute "localsymbol", comboLegObj.ContractSpec.LocalSymbol
        Specifier.setAttribute "multiplier", comboLegObj.ContractSpec.Multiplier
        Specifier.setAttribute "right", comboLegObj.ContractSpec.Right
        Specifier.setAttribute "strike", comboLegObj.ContractSpec.Strike
    Next
End If
ContractToXML = XMLdoc.xml

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateClockFuture( _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pIsSimulated As Boolean, _
                Optional ByVal pClockRate As Single = 1!) As IFuture
Const ProcName As String = "CreateClockFuture"
On Error GoTo Err

Dim lClockFutureBuilder As New ClockFutureBuilder
lClockFutureBuilder.Initialise pContractFuture, pIsSimulated, pClockRate
Set CreateClockFuture = lClockFutureBuilder.Future

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function


Public Function CreateComboLeg( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pIsBuyLeg As Boolean, _
                ByVal pRatio As Long) As ComboLeg
Set CreateComboLeg = New ComboLeg
CreateComboLeg.Initialise pContractSpec, pIsBuyLeg, pRatio
End Function


Public Function CreateContractBuilder( _
                ByVal Specifier As IContractSpecifier) As ContractBuilder
Const ProcName As String = "CreateContractBuilder"
On Error GoTo Err

Set CreateContractBuilder = New ContractBuilder
CreateContractBuilder.Initialise Specifier

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractBuilderFromContract( _
                ByVal pContract As IContract) As ContractBuilder
Const ProcName As String = "CreateContractBuilderFromContract"
On Error GoTo Err

Set CreateContractBuilderFromContract = New ContractBuilder
CreateContractBuilderFromContract.Initialise pContract.Specifier
CreateContractBuilderFromContract.BuildFrom pContract

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractSpecifier( _
                Optional ByVal LocalSymbol As String, _
                Optional ByVal Symbol As String, _
                Optional ByVal TradingClass As String, _
                Optional ByVal Exchange As String, _
                Optional ByVal SecType As SecurityTypes = SecTypeNone, _
                Optional ByVal CurrencyCode As String, _
                Optional ByVal Expiry As String, _
                Optional ByVal Multiplier As Double = 0#, _
                Optional ByVal Strike As Double, _
                Optional ByVal Right As OptionRights = OptNone) As IContractSpecifier
Const ProcName As String = "CreateContractSpecifier"
On Error GoTo Err

Dim lSpec As New ContractSpecifier
lSpec.Initialise LocalSymbol, Symbol, TradingClass, Exchange, SecType, CurrencyCode, Expiry, Multiplier, Strike, Right
Set CreateContractSpecifier = lSpec

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractSpecifierFromString(ByVal pSpecString As String) As IContractSpecifier
Const ProcName As String = "CreateContractSpecifierFromString"
On Error GoTo Err

Const LocalSymbolRegex As String = "(?:\[([a-zA-Z0-9]+(?:(?: *|-|\.)[a-zA-Z0-9]+)*(?:\.[a-zA-Z0-9]?)?)\])?"
Const SecTypeRegEx As String = "(?:([a-zA-Z]+)[\:|#])?"
Const SymbolRegEx As String = "([a-zA-Z0-9]+(?:(?: *|-|\.)[a-zA-Z0-9]+)*(?:\.[a-zA-Z0-9]?)?)(?:/([a-zA-Z0-9]+))?"
Const OptionSpecRegex As String = "(?:=((?:C(?:all)?)|(?:P(?:ut)?))?((?:\d+(?:\.\d{1,2})?))?)?"
Const DateExpiryRegEx As String = "((?:20\d\d)(?:[0|1]\d)(?:[0|1|2|3]\d)?)?"
Const RelativeExpiryRegEx As String = "(\d\d?(?:\[\d\d?d\])?)?"
Const ExpiryRegEx As String = "(?:\(" & _
                                DateExpiryRegEx & _
                                RelativeExpiryRegEx & _
                                "\))?"
Const ExchangeRegEx As String = "(?:@([a-zA-Z]+(?:\-[a-zA-Z]+)?))?"
Const CurrencyRegEx As String = "(?:(?:(?:\$)([a-zA-Z]+))|(?:\(([a-zA-Z]+)\)))?"
Const MultiplierRegex As String = "(?:(?:[\*|'])(\d+(?:\.\d{1,6})?))?"

Const ContractSpecRegex As String = "^" & _
LocalSymbolRegex & _
SecTypeRegEx & _
SymbolRegEx & _
OptionSpecRegex & _
ExpiryRegEx & _
ExchangeRegEx & _
CurrencyRegEx & _
MultiplierRegex & _
"$"


GContracts.RegExpProcessor.Pattern = ContractSpecRegex
GContracts.RegExpProcessor.IgnoreCase = True

Dim lMatches As MatchCollection
Set lMatches = GContracts.RegExpProcessor.Execute(pSpecString)

AssertArgument lMatches.Count = 1, "Invalid contract specifier: [<sectype>:]<symbol>[=<optionspec>][expiry]" & vbCrLf & _
                                   "                            [@<exchange>][(<currency>)][*<multiplier>]" & vbCrLf & _
                                    "Expiry can be: " & vbCrLf & _
                                    "    yyyymm" & vbCrLf & _
                                    "    yyyymmdd" & vbCrLf & _
                                    "    <offset>[<qualifier>d]  for example 0[2d]" & vbCrLf & _
                                    "Option spec: [C[all]|P[ut]][strike]" & vbCrLf & _
                                    "" & vbCrLf & _
                                    "NB: you can use $<currency> instead of (<currency>) if you prefer." & _
                                    "" & vbCrLf & _
                                    "examples: STK:MSFT@SMART$USD" & vbCrLf & _
                                    "          FUT:ESZ0@GLOBEX" & vbCrLf & _
                                    "          FUT:ES(0[2d])@GLOBEX" & vbCrLf & _
                                    "          FUT:ES(202012)@GLOBEX" & vbCrLf & _
                                    "          FUT:DAX(1)@DTB*25" & vbCrLf & _
                                    "          OPT:MSFT=Call285@CBOE"

Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lLocalSymbol As String: lLocalSymbol = lMatch.SubMatches(0)

Dim lSecTypeStr As String: lSecTypeStr = lMatch.SubMatches(1)
Dim lSectype As SecurityTypes

If lSecTypeStr <> "" Then
    lSectype = GContractUtils.SecTypeFromString(lSecTypeStr)
    AssertArgument lSectype <> SecTypeNone, "A valid security type must be supplied"
End If

Dim lSymbolOrLocalSymbol As String
lSymbolOrLocalSymbol = lMatch.SubMatches(2)

Dim lTradingClass As String
lTradingClass = lMatch.SubMatches(3)

Dim lCallPutStr As String: lCallPutStr = lMatch.SubMatches(4)
Dim lCallPut As OptionRights
If lCallPutStr <> "" Then
    If lSectype = SecTypeNone Then
        lSectype = SecTypeOption
    Else
        AssertArgument lSectype = SecTypeOption Or lSectype = SecTypeFuturesOption, _
                        "C(all) or P(ut) can only be supplied for options or futures options"
    End If
    lCallPut = GContractUtils.OptionRightFromString(lCallPutStr)
End If

Dim lStrike As String
lStrike = lMatch.SubMatches(5)
If lStrike = "" Then lStrike = "0"

Dim lDateExpiry As String
lDateExpiry = lMatch.SubMatches(6)

Dim lRelativeExpiry As String
lRelativeExpiry = lMatch.SubMatches(7)

AssertArgument (lDateExpiry = "" And lRelativeExpiry = "") Or _
                lSectype = SecTypeFuture Or _
                lSectype = SecTypeFuturesOption Or _
                lSectype = SecTypeOption Or _
                lSectype = SecTypeNone, _
                lSecTypeStr & " contracts cannot have an expiry specification"
AssertArgument lDateExpiry = "" Or lRelativeExpiry = "", "Supplying both date-based and relative expiry is not permitted"

Dim lExchange As String
lExchange = lMatch.SubMatches(8)

Dim lCurrency As String
lCurrency = lMatch.SubMatches(9)
If lCurrency = "" Then lCurrency = lMatch.SubMatches(10)

Dim lMultiplier As String
lMultiplier = lMatch.SubMatches(11)
If lMultiplier = "" Then lMultiplier = "0"

If (lSectype = SecTypeFuture Or _
    lSectype = SecTypeFuturesOption Or _
    lSectype = SecTypeOption Or _
    lSectype = SecTypeWarrant) _
Then
    If lDateExpiry = "" And lRelativeExpiry = "" Then
        If lTradingClass = "" Then
            Set CreateContractSpecifierFromString = GContractUtils.CreateContractSpecifier( _
                                                            IIf(lLocalSymbol <> "", _
                                                                lLocalSymbol, _
                                                                lSymbolOrLocalSymbol), _
                                                            "", _
                                                            lTradingClass, _
                                                            lExchange, _
                                                            lSectype, _
                                                            lCurrency, _
                                                            "", _
                                                            CDbl(lMultiplier), _
                                                            CDbl(lStrike), _
                                                            lCallPut)
        Else
            Set CreateContractSpecifierFromString = GContractUtils.CreateContractSpecifier( _
                                                            lLocalSymbol, _
                                                            lSymbolOrLocalSymbol, _
                                                            lTradingClass, _
                                                            lExchange, _
                                                            lSectype, _
                                                            lCurrency, _
                                                            "", _
                                                            CDbl(lMultiplier), _
                                                            CDbl(lStrike), _
                                                            lCallPut)
        End If
    Else
        Set CreateContractSpecifierFromString = GContractUtils.CreateContractSpecifier( _
                                                        lLocalSymbol, _
                                                        lSymbolOrLocalSymbol, _
                                                        lTradingClass, _
                                                        lExchange, _
                                                        lSectype, _
                                                        lCurrency, _
                                                        lDateExpiry & lRelativeExpiry, _
                                                        CDbl(lMultiplier), _
                                                        CDbl(lStrike), _
                                                        lCallPut)
    End If
Else
    Set CreateContractSpecifierFromString = GContractUtils.CreateContractSpecifier( _
                                                    lLocalSymbol, _
                                                    lSymbolOrLocalSymbol, _
                                                    lTradingClass, _
                                                    lExchange, _
                                                    lSectype, _
                                                    lCurrency, _
                                                    lDateExpiry & lRelativeExpiry, _
                                                    CDbl(lMultiplier), _
                                                    CDbl(lStrike), _
                                                    lCallPut)
End If

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateContractFromXML( _
                ByVal xmlString As String) As IContract
Const ProcName As String = "CreateContractFromXML"
On Error GoTo Err

Dim lContract As New Contract
lContract.FromXML xmlString
Set CreateContractFromXML = lContract

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSessionBuilderFutureFromContractFuture( _
                ByVal pContractFuture As IFuture, _
                ByVal pUseExchangeTimeZone As Boolean, _
                ByVal pUseFullSession As Boolean) As IFuture
Const ProcName As String = "CreateSessionBuilderFutureFromContractFuture"
On Error GoTo Err

Dim lBuilder As New SessnBuilderFutBldr
lBuilder.Initialise pContractFuture, pUseExchangeTimeZone, pUseFullSession
Set CreateSessionBuilderFutureFromContractFuture = lBuilder.Future

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContracts( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pSingleContractOnly As Boolean, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pListener As IContractFetchListener, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchContracts"
On Error GoTo Err

AssertArgument Not pPrimaryContractStore Is Nothing, "pPrimaryContractStore cannot be Nothing"

If IsEmpty(pCookie) Then pCookie = GenerateGUIDString

Dim t As New ContractFetchTask
t.Initialise pContractSpec, pPrimaryContractStore, pSecondaryContractStore, pCookie, pListener, pSingleContractOnly
StartTask t, pPriority, pTaskName

If pSingleContractOnly Then
    Set FetchContracts = t.ContractFuture
Else
    Set FetchContracts = t.ContractsFuture
End If

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContractsSorted( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                Optional ByVal pSortDescending As Boolean = False, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchContractsSorted"
On Error GoTo Err

AssertArgument Not pPrimaryContractStore Is Nothing, "pPrimaryContractStore cannot be Nothing"

If IsEmpty(pCookie) Then pCookie = GenerateGUIDString

Dim t As New ContractFetchTask
t.InitialiseSorted pContractSpec, _
                    pPrimaryContractStore, _
                    pSecondaryContractStore, _
                    pCookie, _
                    pSortkeys, _
                    pSortDescending
StartTask t, pPriority, pTaskName

Set FetchContractsSorted = t.ContractsFuture

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetExchangeCodes() As String()
Const ProcName As String = "GetExchangeCodes"
On Error GoTo Err

If Not mExchangeCodesInitialised Then setupExchangeCodes
GetExchangeCodes = mExchangeCodes

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchContract( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchContract"
On Error GoTo Err

Set FetchContract = GContractUtils.FetchContracts(pContractSpec, _
                                True, _
                                pPrimaryContractStore, _
                                pSecondaryContractStore, _
                                Nothing, _
                                pPriority, _
                                pTaskName, _
                                pCookie)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchOptionExpiries( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pStrike As Double = 0#, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchOptionExpiries"
On Error GoTo Err

AssertArgument Not pPrimaryContractStore Is Nothing, "pPrimaryContractStore cannot be Nothing"
AssertArgument pPrimaryContractStore.Supports(ContractStoreOptionExpiries), "This Contract Store does not support Option Expiry retrieval"
If Not pSecondaryContractStore Is Nothing Then _
    AssertArgument pSecondaryContractStore.Supports(ContractStoreOptionExpiries), "This Contract Store does not support Option Expiry retrieval"

If IsMissing(pCookie) Then pCookie = GenerateGUIDString

Dim t As New ContractExpsFetchTask
t.Initialise pUnderlyingContractSpecifier, pExchange, pPrimaryContractStore, pStrike, pSecondaryContractStore, pCookie
StartTask t, pPriority, pTaskName

Set FetchOptionExpiries = t.ExpiriesFuture

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FetchOptionStrikes( _
                ByVal pUnderlyingContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pPriority As TaskPriorities = TaskPriorities.PriorityNormal, _
                Optional ByVal pTaskName As String, _
                Optional ByVal pCookie As Variant) As IFuture
Const ProcName As String = "FetchOptionStrikes"
On Error GoTo Err

AssertArgument Not pPrimaryContractStore Is Nothing, "pPrimaryContractStore cannot be Nothing"
AssertArgument pPrimaryContractStore.Supports(ContractStoreOptionStrikes), "This Contract Store does not support Option Strike retrieval"
If Not pSecondaryContractStore Is Nothing Then _
    AssertArgument pSecondaryContractStore.Supports(ContractStoreOptionStrikes), "This Contract Store does not support Option Strike retrieval"

If IsMissing(pCookie) Then pCookie = GenerateGUIDString

Dim t As New ContractStrikesFetchTsk
t.Initialise pUnderlyingContractSpecifier, pExchange, pPrimaryContractStore, pExpiry, pSecondaryContractStore, pCookie
StartTask t, pPriority, pTaskName

Set FetchOptionStrikes = t.StrikesFuture

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPrice( _
                ByVal pPrice As Double, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double) As String
Const ProcName As String = "FormatPrice"
On Error GoTo Err

' see http://www.cmegroup.com/trading/interest-rates/files/TreasuryFuturesPriceRoundingConventions_Mar_24_Final.pdf
' for details of price presentation, especially sections (2) and (7)

FormatPrice = gFormatPrice(pPrice, pSecType, pTickSize)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs32nds( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs32nds"
On Error GoTo Err

FormatPriceAs32nds = gFormatPriceAs32nds(pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs32ndsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs32ndsAndFractions"
On Error GoTo Err

FormatPriceAs32ndsAndFractions = gFormatPriceAs32ndsAndFractions(pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs64ths( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs64ths"
On Error GoTo Err

FormatPriceAs64ths = gFormatPriceAs64ths(pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAs64thsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs64thsAndFractions"
On Error GoTo Err

FormatPriceAs64thsAndFractions = gFormatPriceAs64thsAndFractions(pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function FormatPriceAsDecimals( _
                ByVal pPrice As Double, _
                ByVal pTickSize As Double) As String
Const ProcName As String = "FormatPriceAsDecimals"

On Error GoTo Err

FormatPriceAsDecimals = gFormatPriceAsDecimals(pPrice, pTickSize)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetContractSpecKey(ByVal pSpec As IContractSpecifier) As String
Const ProcName As String = "GetContractSpecKey"
On Error GoTo Err

GetContractSpecKey = pSpec.LocalSymbol & "|" & _
    CStr(pSpec.SecType) & "|" & _
    pSpec.Symbol & "|" & _
    pSpec.TradingClass & "|" & _
    pSpec.Expiry & "|" & _
    pSpec.Strike & "|" & _
    CStr(pSpec.Right) & "|" & _
    pSpec.Exchange & "|" & _
    pSpec.CurrencyCode & "|" & _
    pSpec.Multiplier

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrentQuarterExpiry() As String
Dim year As Long: year = DatePart("yyyy", Now)
Dim quarter As Long: quarter = (Int((DatePart("m", Now) + 2) / 3)) * 3
GetCurrentQuarterExpiry = CStr(year) & Format(quarter, "00")
End Function

Public Function GetNextQuarterExpiry() As String
Dim year As Long: year = DatePart("yyyy", Now)
Dim nextQuarter As Long: nextQuarter = (Int((DatePart("m", Now) + 2) / 3) + 1) * 3
If nextQuarter > 12 Then
    year = year + 1
    nextQuarter = nextQuarter - 12
End If
GetNextQuarterExpiry = CStr(year) & Format(nextQuarter, "00")
End Function

Public Function IsContractExpired( _
                ByVal pContract As IContract, _
                Optional ByVal pClock As Clock) As Boolean
Const ProcName As String = "IsContractExpired"
On Error GoTo Err

Select Case pContract.Specifier.SecType
Case SecTypeFuture, _
        SecTypeOption, _
        SecTypeFuturesOption, _
        SecTypeWarrant
Case Else
    Exit Function
End Select

Dim lExpiry As Date

If Int(pContract.ExpiryDate) <> pContract.ExpiryDate Then
    lExpiry = pContract.ExpiryDate
ElseIf pContract.SessionEndTime <> 0 Then
    lExpiry = pContract.ExpiryDate + pContract.SessionEndTime
Else
    lExpiry = pContract.ExpiryDate + 1
End If

lExpiry = ConvertDateTzToUTC(lExpiry, GetTimeZone(pContract.TimezoneName))

If pClock Is Nothing Then
    IsContractExpired = (lExpiry <= GetTimestampUTC)
Else
    IsContractExpired = (lExpiry <= pClock.TimestampUTC)
End If

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsContractSpecOffsetExpiry( _
                ByVal pContractSpec As IContractSpecifier, _
                Optional ByRef pErrorMessage) As Boolean
Dim lErrorMessage As String
IsContractSpecOffsetExpiry = GContractUtils.IsOffsetExpiry(pContractSpec.Expiry, lErrorMessage)
If Not IsMissing(pErrorMessage) Then pErrorMessage = lErrorMessage
End Function

Public Function IsOffsetExpiry( _
                ByVal Value As String, _
                Optional ByRef ErrorMessage As String) As Boolean
Const ProcName As String = "IsOffsetExpiry"
On Error GoTo Err

If Value = "" Then
    IsOffsetExpiry = False
    Exit Function
End If

Dim l1 As Long
Dim l2 As Long
IsOffsetExpiry = GContractUtils.parseOffsetExp(Value, l1, l2, ErrorMessage)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidExchangeCode(ByVal Code As String) As Boolean
Const ProcName As String = "IsValidExchangeCode"
On Error GoTo Err

If Not mExchangeCodesInitialised Then setupExchangeCodes

Code = UCase$(Code)

IsValidExchangeCode = BinarySearchStrings( _
                            Code, _
                            mExchangeCodes, _
                            IsCaseSensitive:=False) >= 0

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidExchangeSpecifier(ByVal pExchangeSpec As String) As Boolean
Const ProcName As String = "IsValidExchangeSpecifier"
On Error GoTo Err

pExchangeSpec = UCase$(pExchangeSpec)

Const ExchangeSmartQualified As String = "SMART-"
If InStr(1, pExchangeSpec, ExchangeSmartQualified) <> 1 Then
    IsValidExchangeSpecifier = GContractUtils.IsValidExchangeCode(pExchangeSpec)
    Exit Function
End If

IsValidExchangeSpecifier = GContractUtils.IsValidExchangeCode( _
                                Right$(pExchangeSpec, _
                                        Len(pExchangeSpec) - Len(ExchangeSmartQualified)))

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidExpiry( _
                ByVal Value As String, _
                Optional ByRef ErrorMessage As String) As Boolean
Const ProcName As String = "IsValidExpiry"
On Error GoTo Err

If Value = "" Then
    IsValidExpiry = True
    Exit Function
End If

If GContractUtils.IsOffsetExpiry(Value, ErrorMessage) Then
    IsValidExpiry = True
    Exit Function
End If

Dim d As Date

If IsDate(Value) Then
    d = CDate(Value)
ElseIf Len(Value) = 8 Then
    Dim datestring As String
    datestring = Left$(Value, 4) & "/" & Mid$(Value, 5, 2) & "/" & Right$(Value, 2)
    If IsDate(datestring) Then d = CDate(datestring)
End If

If d <> 0 Then
    If d >= CDate("2000/01/01") And d <= CDate((year(Now) + 10) & "/12/31") Then
        IsValidExpiry = True
        Exit Function
    End If
End If

If Len(Value) = 6 Then
    If IsInteger(Value, 200001, (year(Now) + 10) * 100 + 12) Then
        If Right$(Value, 2) <= 12 Then
            IsValidExpiry = True
            Exit Function
        End If
    End If
End If

IsValidExpiry = False

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidPrice( _
                ByVal pPrice As Double, _
                ByVal pPrevValidPrice As Double, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double) As Boolean
Const ProcName As String = "IsValidPrice"
On Error GoTo Err

If pTickSize = 0 Then
    IsValidPrice = True
    Exit Function
End If

If pPrevValidPrice = 0 Or pPrevValidPrice = MaxDouble Then
    If Abs(pPrice) / pTickSize > &H7FFFFF Then Exit Function ' note that Z index has ticksize 0.01
                                                            ' so we need to allow plenty of room
                                                            ' &H7FFFFF = 8388607
    
    ' A first price of 0 is always considered invalid. Although some indexes can validly
    ' have zero values, it is unlikely on the very first price notified, and this check
    ' catches the occasional zero prices sent by IB when, for example, the Z ticker
    ' is started
    If pPrice = 0 Then Exit Function
    
    If pSecType = SecTypeIndex Then
        ' don't do this check for indexes because some of them, such as TICK-NYSE, can have both
        ' positive, zero and negative values
    Else
        If pPrice < 0 Then Exit Function
        If pPrice / pTickSize < 1 Then Exit Function ' catch occasional very small prices from IB
    End If
Else
    'If Abs(pPrevValidPrice - pPrice) / pTickSize > 32767 Then Exit Function
    
    If pSecType = SecTypeIndex Then
        ' don't do this check for indexes because some of them, such as TICK-NYSE, can have both
        ' positive and negative values - moreover the value can change dramatically from one
        ' tick to the next
    Else
        If pPrice <= 0 Then Exit Function
        If pPrice / pTickSize < 1 Then Exit Function ' catch occasional very small prices from IB
        'If pPrice < (2 * pPrevValidPrice) / 3 Or pPrice > (3 * pPrevValidPrice) / 2 Then Exit Function
    End If
    
End If

IsValidPrice = True

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidSecType( _
                ByVal Value As Long) As Boolean
IsValidSecType = True
Select Case Value
Case SecTypeStock
Case SecTypeFuture
Case SecTypeOption
Case SecTypeFuturesOption
Case SecTypeCash
Case SecTypeCombo
Case SecTypeIndex
Case SecTypeWarrant
Case Else
    IsValidSecType = False
End Select
End Function

Public Function LoadContractFromConfig(ByVal pConfig As ConfigurationSection) As IContract
Const ProcName As String = "LoadContractFromConfig"
On Error GoTo Err

Dim lSpec As ContractSpecifier
Set lSpec = LoadContractSpecFromConfig(pConfig.AddConfigurationSection(ConfigSectionContractSpecifier))

Dim lContract As New Contract
lContract.Specifier = lSpec

With pConfig
    lContract.DaysBeforeExpiryToSwitch = .GetSetting(ConfigSettingDaysBeforeExpiryToSwitch, DefaultDaysBeforeExpiryToSwitch)
    lContract.Description = .GetSetting(ConfigSettingDescription, "")
    lContract.ExpiryDate = CDate(.GetSetting(ConfigSettingExpiryDate, DefaultExpiry))
    lContract.FullSessionEndTime = .GetSetting(ConfigSettingFullSessionEndTime, "00:00:00")
    lContract.FullSessionStartTime = .GetSetting(ConfigSettingFullSessionStartTime, "00:00:00")
    lContract.SessionEndTime = .GetSetting(ConfigSettingSessionEndTime, "00:00:00")
    lContract.SessionStartTime = .GetSetting(ConfigSettingSessionStartTime, "00:00:00")
    lContract.TickSize = .GetSetting(ConfigSettingTickSize, DefaultTickSize)
    lContract.TimezoneName = .GetSetting(ConfigSettingTimezoneName, DefaultTimezoneName)
    If .GetSetting(ConfigSettingMultiplier, DefaultMultiplier) <> DefaultMultiplier Then
        lSpec.Multiplier = .GetSetting(ConfigSettingMultiplier, DefaultMultiplier)
    End If
End With

Set LoadContractFromConfig = lContract

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadContractSpecFromConfig(ByVal pConfig As ConfigurationSection) As ContractSpecifier
Const ProcName As String = "LoadContractSpecFromConfig"
On Error GoTo Err

Dim lContractSpec As ContractSpecifier
With pConfig
    Set lContractSpec = GContractUtils.CreateContractSpecifier(.GetSetting(ConfigSettingContractSpecLocalSymbol, ""), _
                                            .GetSetting(ConfigSettingContractSpecSymbol, ""), _
                                            .GetSetting(ConfigSettingContractSpecTradingClass, ""), _
                                            .GetSetting(ConfigSettingContractSpecExchange, ""), _
                                            GContractUtils.SecTypeFromString(.GetSetting(ConfigSettingContractSpecSecType, "")), _
                                            .GetSetting(ConfigSettingContractSpecCurrency, ""), _
                                            .GetSetting(ConfigSettingContractSpecExpiry, ""), _
                                            CDbl(.GetSetting(ConfigSettingContractSpecMultiplier, DefaultMultiplier)), _
                                            CDbl(.GetSetting(ConfigSettingContractSpecStrikePrice, "0.0")), _
                                            GContractUtils.OptionRightFromString(.GetSetting(ConfigSettingContractSpecRight, "")))
End With

Set LoadContractSpecFromConfig = lContractSpec

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ParseOffsetExpiry( _
                ByVal Value As String, _
                ByRef pExpiryOffset As Long, _
                ByRef pDaysBeforeExpiryToSwitch As Long)
                
Const ProcName As String = "ParseOffsetExpiry"
On Error GoTo Err

Dim lErrorMessage As String
If Not GContractUtils.parseOffsetExp(Value, _
                            pExpiryOffset, _
                            pDaysBeforeExpiryToSwitch, _
                            lErrorMessage) _
Then
    AssertArgument False, lErrorMessage
End If

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ParsePrice( _
                ByVal pPriceString As String, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePrice"
On Error GoTo Err

ParsePrice = gParsePrice(pPriceString, pSecType, pTickSize, pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs32nds( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs32nds"
On Error GoTo Err

ParsePriceAs32nds = gParsePriceAs32nds(pPriceString, pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs32ndsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs32ndsAndFractions"
On Error GoTo Err

ParsePriceAs32ndsAndFractions = gParsePriceAs32ndsAndFractions(pPriceString, pPrice)
    
Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs64ths( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs64ths"
On Error GoTo Err

ParsePriceAs64ths = gParsePriceAs64ths(pPriceString, pPrice)
    
Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAs64thsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs64thsAndFractions"
On Error GoTo Err

ParsePriceAs64thsAndFractions = gParsePriceAs64thsAndFractions(pPriceString, pPrice)
    
Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAsDecimals( _
                ByVal pPriceString As String, _
                ByVal pTickSize As Double, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAsDecimals"
On Error GoTo Err

ParsePriceAsDecimals = gParsePriceAsDecimals(pPriceString, pTickSize, pPrice)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SaveContractSpecToConfig(ByVal pContractSpec As IContractSpecifier, ByVal pConfig As ConfigurationSection)
Const ProcName As String = "SaveContractSpecToConfig"
On Error GoTo Err

With pConfig
    .SetSetting ConfigSettingContractSpecLocalSymbol, pContractSpec.LocalSymbol
    .SetSetting ConfigSettingContractSpecSymbol, pContractSpec.Symbol
    .SetSetting ConfigSettingContractSpecTradingClass, pContractSpec.TradingClass
    .SetSetting ConfigSettingContractSpecExchange, pContractSpec.Exchange
    .SetSetting ConfigSettingContractSpecSecType, GContractUtils.SecTypeToString(pContractSpec.SecType)
    .SetSetting ConfigSettingContractSpecCurrency, pContractSpec.CurrencyCode
    .SetSetting ConfigSettingContractSpecExpiry, pContractSpec.Expiry
    .SetSetting ConfigSettingContractSpecMultiplier, pContractSpec.Multiplier
    .SetSetting ConfigSettingContractSpecStrikePrice, pContractSpec.Strike
    .SetSetting ConfigSettingContractSpecRight, GContractUtils.OptionRightToString(pContractSpec.Right)
End With

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SaveContractToConfig(ByVal pContract As IContract, ByVal pConfig As ConfigurationSection)
Const ProcName As String = "SaveContractToConfig"
On Error GoTo Err

SaveContractSpecToConfig pContract.Specifier, pConfig.AddConfigurationSection(ConfigSectionContractSpecifier)

With pConfig
    .SetSetting ConfigSettingDaysBeforeExpiryToSwitch, pContract.DaysBeforeExpiryToSwitch
    .SetSetting ConfigSettingDescription, pContract.Description
    .SetSetting ConfigSettingExpiryDate, FormatTimestamp(pContract.ExpiryDate, TimestampDateOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingSessionEndTime, FormatTimestamp(pContract.SessionEndTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingSessionStartTime, FormatTimestamp(pContract.SessionStartTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingFullSessionEndTime, FormatTimestamp(pContract.FullSessionEndTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingFullSessionStartTime, FormatTimestamp(pContract.FullSessionStartTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingTickSize, pContract.TickSize
    .SetSetting ConfigSettingTimezoneName, pContract.TimezoneName
End With

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SecTypeFromString(ByVal Value As String) As SecurityTypes
Select Case UCase$(Value)
Case "STOCK", "STK"
    SecTypeFromString = SecTypeStock
Case "FUTURE", "FUT"
    SecTypeFromString = SecTypeFuture
Case "OPTION", "OPT"
    SecTypeFromString = SecTypeOption
Case "FUTURES OPTION", "FOP"
    SecTypeFromString = SecTypeFuturesOption
Case "CASH"
    SecTypeFromString = SecTypeCash
Case "COMBO", "CMB"
    SecTypeFromString = SecTypeCombo
Case "INDEX", "IND"
    SecTypeFromString = SecTypeIndex
Case "WARRANT", "WAR"
    SecTypeFromString = SecTypeWarrant
Case "CFD"
    SecTypeFromString = SecTypeCFD
Case "CRYPTO CURRENCY", "CRYPTOCURRENCY", "CRYPTO"
    SecTypeFromString = SecTypeCrypto
Case Else
    SecTypeFromString = SecTypeNone
End Select
End Function

Public Function SecTypeToString(ByVal Value As SecurityTypes) As String
Select Case Value
Case SecTypeStock
    SecTypeToString = "Stock"
Case SecTypeFuture
    SecTypeToString = "Future"
Case SecTypeOption
    SecTypeToString = "Option"
Case SecTypeFuturesOption
    SecTypeToString = "Futures Option"
Case SecTypeCash
    SecTypeToString = "Cash"
Case SecTypeCombo
    SecTypeToString = "Combo"
Case SecTypeIndex
    SecTypeToString = "Index"
Case SecTypeWarrant
    SecTypeToString = "Warrant"
Case SecTypeCFD
    SecTypeToString = "CFD"
Case SecTypeCrypto
    SecTypeToString = "Crypto Currency"
End Select
End Function

Public Function SecTypeToShortString(ByVal Value As SecurityTypes) As String
Select Case Value
Case SecTypeStock
    SecTypeToShortString = "STK"
Case SecTypeFuture
    SecTypeToShortString = "FUT"
Case SecTypeOption
    SecTypeToShortString = "OPT"
Case SecTypeFuturesOption
    SecTypeToShortString = "FOP"
Case SecTypeCash
    SecTypeToShortString = "CASH"
Case SecTypeCombo
    SecTypeToShortString = "CMB"
Case SecTypeIndex
    SecTypeToShortString = "IND"
Case SecTypeWarrant
    SecTypeToShortString = "WAR"
Case SecTypeCFD
    SecTypeToShortString = "CFD"
Case SecTypeCrypto
    SecTypeToShortString = "CRYPTO"
End Select
End Function

Public Function TryCreateContractSpecifierFromString( _
                ByVal SpecString As String, _
                ByRef ContractSpec As IContractSpecifier, _
                Optional ByRef ErrorMessage As String) As Boolean
Const ProcName As String = "CreateContractSpecifierFromString"
On Error GoTo Err


Set ContractSpec = GContractUtils.CreateContractSpecifierFromString(SpecString)

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    ErrorMessage = Err.Description
    TryCreateContractSpecifierFromString = False
Else
    GContracts.HandleUnexpectedError ProcName, ModuleName
End If
End Function

Public Function TryParseOffsetExpiry( _
                ByVal Value As String, _
                ByRef pExpiryOffset As Long, _
                ByRef pDaysBeforeExpiryToSwitch As Long, _
                Optional ByRef pErrorMessage As String) As Boolean
Const ProcName As String = "TryParseOffsetExpiry"
On Error GoTo Err

TryParseOffsetExpiry = GContractUtils.parseOffsetExp(Value, _
                                    pExpiryOffset, _
                                    pDaysBeforeExpiryToSwitch, _
                                    pErrorMessage)
                                    
Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addExchangeCode(ByVal Code As String)
Const ProcName As String = "addExchangeCode"
On Error GoTo Err

mMaxExchangeCodesIndex = mMaxExchangeCodesIndex + 1
If mMaxExchangeCodesIndex > UBound(mExchangeCodes) Then
    ReDim Preserve mExchangeCodes(2 * (UBound(mExchangeCodes) + 1) - 1) As String
End If
mExchangeCodes(mMaxExchangeCodesIndex) = UCase$(Code)

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function OptionRightFromString(ByVal Value As String) As OptionRights
Select Case UCase$(Value)
Case ""
    OptionRightFromString = OptNone
Case "CALL", "C"
    OptionRightFromString = OptCall
Case "PUT", "P"
    OptionRightFromString = OptPut
End Select
End Function

Public Function OptionRightToShortString(ByVal Value As OptionRights) As String
Select Case Value
Case OptNone
    OptionRightToShortString = ""
Case OptCall
    OptionRightToShortString = "C"
Case OptPut
    OptionRightToShortString = "P"
End Select
End Function

Public Function OptionRightToString(ByVal Value As OptionRights) As String
Select Case Value
Case OptNone
    OptionRightToString = ""
Case OptCall
    OptionRightToString = "Call"
Case OptPut
    OptionRightToString = "Put"
End Select
End Function

Private Function parseOffsetExp( _
                ByVal Value As String, _
                ByRef pExpiryOffset As Long, _
                ByRef pDaysBeforeExpiryToSwitch As Long, _
                ByRef pMessage As String) As Boolean
Const ProcName As String = "ParseOffsetExp"
On Error GoTo Err

Const OffsetExpiryFormat As String = "^(\d\d?)(?:\[(\d\d?)d\])?$"

GContracts.RegExpProcessor.Pattern = OffsetExpiryFormat
GContracts.RegExpProcessor.IgnoreCase = True

Dim lMatches As MatchCollection
Set lMatches = GContracts.RegExpProcessor.Execute(Trim$(Value))

If Not lMatches.Count = 1 Then
    pMessage = "Expiry syntax invalid"
    parseOffsetExp = False
    Exit Function
End If

Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lOffsetStr As String
lOffsetStr = lMatch.SubMatches(0)
If lOffsetStr = "" Then
    pExpiryOffset = 0
ElseIf IsInteger(lOffsetStr, 0, MaxContractExpiryOffset) Then
    pExpiryOffset = CInt(lOffsetStr)
Else
    pMessage = "Expiry offset must be >= 0 and <= " & MaxContractExpiryOffset
    parseOffsetExp = False
    Exit Function
End If

Dim lDaysBeforeExpiryStr As String
lDaysBeforeExpiryStr = lMatch.SubMatches(1)
If lDaysBeforeExpiryStr = "" Then
    pDaysBeforeExpiryToSwitch = 0
ElseIf IsInteger(lDaysBeforeExpiryStr, 0, MaxContractDaysBeforeExpiryToSwitch) Then
    pDaysBeforeExpiryToSwitch = CInt(lDaysBeforeExpiryStr)
Else
    pMessage = "Expiry modifier must be >= 0 and <= " & MaxContractDaysBeforeExpiryToSwitch
    parseOffsetExp = False
    Exit Function
End If

parseOffsetExp = True

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setupExchangeCodes()
Const ProcName As String = "setupExchangeCodes"
On Error GoTo Err

ReDim mExchangeCodes(31) As String
mMaxExchangeCodesIndex = -1

addExchangeCode "ACE"
addExchangeCode "AEB"
addExchangeCode "AEQLIT"
addExchangeCode "AMEX"
addExchangeCode "AQXEUK"
addExchangeCode "ARCA"
addExchangeCode "ASX"

addExchangeCode "BATEUK"
addExchangeCode "BATS"
addExchangeCode "BELFOX"
addExchangeCode "BEX"
addExchangeCode "BOX"
addExchangeCode "BRUT"
addExchangeCode "BTRADE"
addExchangeCode "BVME"
addExchangeCode "BYX"

addExchangeCode "CAES"
addExchangeCode "CBOE"
addExchangeCode "CBOE2"
addExchangeCode "CBOT"
addExchangeCode "CBSX"
addExchangeCode "CDE"
addExchangeCode "CFE"
addExchangeCode "CHIXUK"
addExchangeCode "CHX"
addExchangeCode "CME"
addExchangeCode "COMEX"
addExchangeCode "CSFBALGO"

addExchangeCode "DRCTEDGE"
addExchangeCode "DTB"

addExchangeCode "EBS"
addExchangeCode "ECBOT"
addExchangeCode "EDGEA"
addExchangeCode "EDGX"
addExchangeCode "EMERALD"
addExchangeCode "EUIBSI"
addExchangeCode "EUREX"
addExchangeCode "EUREXUS"

addExchangeCode "FOXRIVER"
addExchangeCode "FTA"
addExchangeCode "FWB"
addExchangeCode "FWB2"

addExchangeCode "GEMINI"
addExchangeCode "GLOBEX"
addExchangeCode "GETTEX2"

addExchangeCode "HKFE"

addExchangeCode "IBEOS"
addExchangeCode "IBIS"
addExchangeCode "ICEEU"
addExchangeCode "ICEEUSOFT"
addExchangeCode "IDEAL"
addExchangeCode "IDEALPRO"
addExchangeCode "IDEM"
addExchangeCode "IEX"
addExchangeCode "INET"
addExchangeCode "INSTINET"
addExchangeCode "ISE"
addExchangeCode "ISLAND"

addExchangeCode "JEFFALGO"

addExchangeCode "LAVA"
addExchangeCode "LIFFE"
addExchangeCode "LIFFE_NF"
addExchangeCode "LSE"
addExchangeCode "LSEETF"
addExchangeCode "LSEIOB1"
addExchangeCode "LSSF"
addExchangeCode "LTSE"

addExchangeCode "MATIF"
addExchangeCode "MEFF"
addExchangeCode "MEFFRV"
addExchangeCode "MEMX"
addExchangeCode "MERCURY"
addExchangeCode "MEXI"
addExchangeCode "MIAX"
addExchangeCode "MONEP"
addExchangeCode "MXT"

addExchangeCode "NASDAQ"
addExchangeCode "NASDAQBX"
addExchangeCode "NASDAQOM"
addExchangeCode "NQLX"
addExchangeCode "NSE"
addExchangeCode "NSX"
addExchangeCode "NYBOT"
addExchangeCode "NYMEX"
addExchangeCode "NYSE"
addExchangeCode "NYSENAT"

addExchangeCode "OMS"
addExchangeCode "OMXNO"
addExchangeCode "ONE"
addExchangeCode "OSE"
addExchangeCode "OSE.JPN"
addExchangeCode "OVERNIGHT"

addExchangeCode "PAXOS"
addExchangeCode "PEARL"
addExchangeCode "PHLX"
addExchangeCode "PINK"
addExchangeCode "PSE"
addExchangeCode "PSX"
addExchangeCode "PURE"

addExchangeCode "QBALGO"

addExchangeCode "RDBK"

addExchangeCode "SBF"
addExchangeCode "SEHK"
addExchangeCode "SFB"
addExchangeCode "SGX"
addExchangeCode "SMART"
addExchangeCode "SMARTCAN"
addExchangeCode "SMARTEUR"
addExchangeCode "SMARTNASDAQ"
addExchangeCode "SMARTNYSE"
addExchangeCode "SMARTUK"
addExchangeCode "SMARTUS"
addExchangeCode "SNFE"
addExchangeCode "SOFFEX"
addExchangeCode "SWB"
addExchangeCode "SWB2"
addExchangeCode "SWX"

addExchangeCode "TPLUS1"
addExchangeCode "TPLUS2"
addExchangeCode "TPLUS0"
addExchangeCode "TRACKECN"
addExchangeCode "TRQXEN"
addExchangeCode "TRQXUK"
addExchangeCode "TSE"
addExchangeCode "TSE.JPN"

addExchangeCode "VALUE"
addExchangeCode "VENTURE"
addExchangeCode "VIRTX"
addExchangeCode "VSE"
addExchangeCode "VWAP"

ReDim Preserve mExchangeCodes(mMaxExchangeCodesIndex) As String
SortStrings mExchangeCodes

mExchangeCodesInitialised = True

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function UnEscapeRegexSpecialChar( _
                ByRef inString As String) As String
Dim s As String
Dim skipNextCheck As Boolean
Dim i As Long
For i = 1 To Len(inString)
    Dim ch As String: ch = Mid$(inString, i, 1)
    If skipNextCheck Then
        s = s & ch
        skipNextCheck = False
    Else
        If ch = "\" Then
            skipNextCheck = True
        Else
            s = s & ch
        End If
    End If
Next

UnEscapeRegexSpecialChar = s
End Function





