VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements Comparable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "Contract"

Private Const OneThirtySecond           As Double = 0.03125
Private Const OneSixtyFourth            As Double = 0.015625
Private Const OneHundredTwentyEighth    As Double = 0.0078125

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSpecifier As ContractSpecifier
Private mTickSize As Double
Private mMultiplier As Double
Private mNumberOfDecimals As Long
Private mDescription As String
Private mSessionStartTime As Date
Private mSessionEndTime As Date
Private mExpiryDate As Date
Private mDaysBeforeExpiryToSwitch As Long
Private mProviderIDs As Parameters
Private mTimeZone As Timezone

Private mSortKeys() As ContractSortKeyIds

Private mPriceFormatter As PriceFormatter
Private mPriceParser As PriceParser

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPriceFormatter = New PriceFormatter
Set mPriceParser = New PriceParser
End Sub

'@================================================================================
' Comparable Interface Members
'@================================================================================

Private Function Comparable_CompareTo( _
                ByVal obj As Comparable) As Long
Dim i As Long
Dim other As Contract
Dim otherSpec As ContractSpecifier

Const ProcName As String = "Comparable_CompareTo"
Dim failpoint As String
On Error GoTo Err

Set other = obj
Set otherSpec = other.Specifier

For i = 0 To UBound(mSortKeys)
    Select Case mSortKeys(i)
    Case ContractSortKeyNone
        Exit Function
    Case ContractSortKeyLocalSymbol
        Comparable_CompareTo = StrComp(mSpecifier.LocalSymbol, otherSpec.LocalSymbol, vbTextCompare)
    Case ContractSortKeySymbol
        Comparable_CompareTo = StrComp(mSpecifier.Symbol, otherSpec.Symbol, vbTextCompare)
    Case ContractSortKeySecType
        Comparable_CompareTo = StrComp(gSecTypeToShortString(mSpecifier.SecType), gSecTypeToShortString(otherSpec.SecType), vbTextCompare)
    Case ContractSortKeyExchange
        Comparable_CompareTo = StrComp(mSpecifier.Exchange, otherSpec.Exchange, vbTextCompare)
    Case ContractSortKeyExpiry
        Comparable_CompareTo = StrComp(mSpecifier.Expiry, otherSpec.Expiry, vbTextCompare)
    Case ContractSortKeyCurrency
        Comparable_CompareTo = StrComp(mSpecifier.CurrencyCode, otherSpec.CurrencyCode, vbTextCompare)
    Case ContractSortKeyRight
        Comparable_CompareTo = StrComp(gOptionRightToString(mSpecifier.Right), gOptionRightToString(otherSpec.Right), vbTextCompare)
    Case ContractSortKeyStrike
        Comparable_CompareTo = Sgn(mSpecifier.Strike - otherSpec.Strike)
    End Select
    If Comparable_CompareTo <> 0 Then Exit Function
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Properties
'@================================================================================

Public Property Get DaysBeforeExpiryToSwitch() As Long
DaysBeforeExpiryToSwitch = mDaysBeforeExpiryToSwitch
End Property

Friend Property Let DaysBeforeExpiryToSwitch(ByVal value As Long)
mDaysBeforeExpiryToSwitch = value
End Property

Friend Property Let Description(ByVal value As String)
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Get ExpiryDate() As Date
ExpiryDate = mExpiryDate
End Property

Friend Property Let ExpiryDate(ByVal value As Date)
mExpiryDate = value
End Property

Public Property Get Key() As String
Const ProcName As String = "Key"
Dim failpoint As String
On Error GoTo Err

Key = Specifier.Key

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Multiplier() As Double
Multiplier = mMultiplier
End Property

Friend Property Let Multiplier(ByVal value As Double)
mMultiplier = value
End Property

Public Property Get NumberOfDecimals() As Long
NumberOfDecimals = mNumberOfDecimals
End Property

Public Property Get ProviderID(ByVal providerKey As String) As String
Const ProcName As String = "ProviderID"
Dim failpoint As String
On Error GoTo Err

ProviderID = mProviderIDs.GetParameterValue(providerKey, mSpecifier.LocalSymbol)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let ProviderIDs(ByVal value As Parameters)
Set mProviderIDs = value
End Property

Friend Property Get ProviderIDs() As Parameters
Set ProviderIDs = mProviderIDs
End Property

Friend Property Let SessionStartTime(ByVal value As Date)
mSessionStartTime = TimeValue(value)
End Property

Public Property Get SessionStartTime() As Date
SessionStartTime = mSessionStartTime
End Property

Friend Property Let SessionEndTime(ByVal value As Date)
mSessionEndTime = TimeValue(value)
End Property

Public Property Get SessionEndTime() As Date
SessionEndTime = mSessionEndTime
End Property

Friend Property Let SortKeys( _
                ByRef value() As ContractSortKeyIds)
mSortKeys = value
End Property
                
Public Property Get Specifier() As ContractSpecifier
Set Specifier = mSpecifier
End Property

Friend Property Let Specifier(ByVal value As ContractSpecifier)
Set mSpecifier = value
End Property

Public Property Get TickSize() As Double
TickSize = mTickSize
End Property

Friend Property Let TickSize(ByVal value As Double)
Dim minTickString As String

Const ProcName As String = "TickSize"
Dim failpoint As String
On Error GoTo Err

mTickSize = value

minTickString = Format(mTickSize, "0.##############")

mNumberOfDecimals = Len(minTickString) - 2

mPriceParser.NumberOfDecimals = mNumberOfDecimals

If mNumberOfDecimals = 0 Then
    mPriceFormatter.PriceFormatString = "0"
Else
    mPriceFormatter.PriceFormatString = "0." & String(mNumberOfDecimals, "0")
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Property

Public Property Get TickValue() As Double
TickValue = mTickSize * mMultiplier
End Property

Public Property Get Timezone() As Timezone
Set Timezone = mTimeZone
End Property

Friend Property Let Timezone( _
                ByVal value As Timezone)
Set mTimeZone = value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function FormatPrice( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPrice"
Dim failpoint As String
On Error GoTo Err

' see http://www.cmegroup.com/trading/interest-rates/files/TreasuryFuturesPriceRoundingConventions_Mar_24_Final.pdf
' for details of price presentation, especially sections (2) and (7)

If mTickSize = OneThirtySecond Then
    FormatPrice = FormatPriceAs32nds(pPrice)
ElseIf mTickSize = OneSixtyFourth Then
    If mSpecifier.SecType = SecTypeFuture Then
        FormatPrice = FormatPriceAs32ndsAndFractions(pPrice)
    Else
        FormatPrice = FormatPriceAs64ths(pPrice)
    End If
ElseIf mTickSize = OneHundredTwentyEighth Then
    If mSpecifier.SecType = SecTypeFuture Then
        FormatPrice = FormatPriceAs32ndsAndFractions(pPrice)
    Else
        FormatPrice = FormatPriceAs64thsAndFractions(pPrice)
    End If
Else
    FormatPrice = FormatPriceAsDecimals(pPrice)
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAs32nds( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs32nds"
Dim failpoint As String
On Error GoTo Err

FormatPriceAs32nds = mPriceFormatter.FormatPriceAs32nds(pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAs32ndsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs32ndsAndFractions"
Dim failpoint As String
On Error GoTo Err

FormatPriceAs32ndsAndFractions = mPriceFormatter.FormatPriceAs32ndsAndFractions(pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAs64ths( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs64ths"
Dim failpoint As String
On Error GoTo Err

FormatPriceAs64ths = mPriceFormatter.FormatPriceAs64ths(pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAs64thsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAs64thsAndFractions"
Dim failpoint As String
On Error GoTo Err

FormatPriceAs64thsAndFractions = mPriceFormatter.FormatPriceAs64thsAndFractions(pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAsDecimals( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAsDecimals"
Dim failpoint As String
On Error GoTo Err

FormatPriceAsDecimals = mPriceFormatter.FormatPriceAsDecimals(pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Sub FromXML(ByVal contractXML As String)
Dim XMLdoc As DOMDocument30
Dim contractEl As IXMLDOMElement
Dim specifierEl As IXMLDOMElement
Dim specifierObj As ContractSpecifier
Dim comboLegsEl As IXMLDOMElement
Dim comboLegEl As IXMLDOMElement
Dim comboLegObj As ComboLeg
Dim nodeList As IXMLDOMNodeList
Dim comboSpecifierEl As IXMLDOMElement
Dim comboSpecifierObj As ContractSpecifier

Const ProcName As String = "FromXML"
Dim failpoint As String
On Error GoTo Err

Set XMLdoc = New DOMDocument30
XMLdoc.loadXML contractXML
If XMLdoc.parseError.errorCode <> 0 Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Contract XML cannot be parsed"
End If

Set contractEl = XMLdoc.selectSingleNode("contract")
TickSize = contractEl.getAttribute("minimumtick")
mMultiplier = contractEl.getAttribute("multiplier")
If IsNull(contractEl.getAttribute("sessionstarttime")) Then
    mSessionStartTime = "00:00:00"
Else
    mSessionStartTime = contractEl.getAttribute("sessionstarttime")
End If
If IsNull(contractEl.getAttribute("sessionendtime")) Then
    mSessionEndTime = "00:00:00"
Else
    mSessionEndTime = contractEl.getAttribute("sessionendtime")
End If
If IsNull(contractEl.getAttribute("description")) Then
    mDescription = ""
Else
    mDescription = contractEl.getAttribute("description")
End If
If IsNull(contractEl.getAttribute("numberofdecimals")) Then
    mNumberOfDecimals = 4
Else
    mNumberOfDecimals = contractEl.getAttribute("numberofdecimals")
End If
If IsNull(contractEl.getAttribute("timezonename")) Then
    Set mTimeZone = GetTimeZone()
Else
    Set mTimeZone = GetTimeZone(CStr(contractEl.getAttribute("timezonename")))
End If

Set specifierEl = contractEl.selectSingleNode("specifier")
Set specifierObj = New ContractSpecifier
Me.Specifier = specifierObj
With specifierObj
    .CurrencyCode = specifierEl.getAttribute("currencycode")
    .Exchange = specifierEl.getAttribute("exchange")
    .Expiry = specifierEl.getAttribute("expiry")
    .LocalSymbol = specifierEl.getAttribute("localsymbol")
    .Right = specifierEl.getAttribute("right")
    .SecType = specifierEl.getAttribute("sectype")
    .Strike = specifierEl.getAttribute("strike")
    .Symbol = specifierEl.getAttribute("symbol")

    Set comboLegsEl = specifierEl.selectSingleNode("combolegs")
    Set nodeList = comboLegsEl.selectNodes("comboleg")
    If nodeList.length <> 0 Then .ComboLegs = New ComboLegs
    For Each comboLegEl In nodeList
        Set comboLegObj = New ComboLeg
        .ComboLegs.Add comboLegObj
        Set comboSpecifierEl = comboLegEl.selectSingleNode("specifier")
        Set comboSpecifierObj = New ContractSpecifier
        With comboSpecifierObj
            .CurrencyCode = comboSpecifierEl.getAttribute("currencycode")
            .Exchange = comboSpecifierEl.getAttribute("exchange")
            .Expiry = comboSpecifierEl.getAttribute("expiry")
            .LocalSymbol = comboSpecifierEl.getAttribute("localsymbol")
            .Right = comboSpecifierEl.getAttribute("right")
            .SecType = comboSpecifierEl.getAttribute("sectype")
            .Strike = comboSpecifierEl.getAttribute("strike")
            .Symbol = comboSpecifierEl.getAttribute("symbol")
        End With
        comboLegObj.Initialise comboSpecifierObj, _
                                CBool(comboLegEl.getAttribute("isbuyleg")), _
                                CLng(comboLegEl.getAttribute("Ratio"))
    Next

End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function IsValidPrice( _
                ByVal pPrice As Double, _
                ByVal pPrevValidPrice As Double) As Boolean
Const ProcName As String = "IsValidPrice"
Dim failpoint As String
On Error GoTo Err

IsValidPrice = gIsValidPrice(pPrice, pPrevValidPrice, mSpecifier.SecType, mTickSize)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function ParsePrice( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePrice"
Dim failpoint As String
On Error GoTo Err

pPriceString = Trim$(pPriceString)

If mTickSize = OneThirtySecond Then
    ParsePrice = ParsePriceAs32nds(pPriceString, pPrice)
ElseIf mTickSize = OneSixtyFourth Then
    If mSpecifier.SecType = SecTypeFuture Then
        ParsePrice = ParsePriceAs32ndsAndFractions(pPriceString, pPrice)
    Else
        ParsePrice = ParsePriceAs64ths(pPriceString, pPrice)
    End If
ElseIf mTickSize = OneHundredTwentyEighth Then
    If mSpecifier.SecType = SecTypeFuture Then
        ParsePrice = ParsePriceAs32ndsAndFractions(pPriceString, pPrice)
    Else
        ParsePrice = ParsePriceAs64thsAndFractions(pPriceString, pPrice)
    End If
Else
    ParsePrice = ParsePriceAsDecimals(pPriceString, pPrice)
End If

If pPrice <= 0 Then Exit Function

If Abs(pPrice - mTickSize * Int((pPrice + mTickSize / 1000) / mTickSize)) > (mTickSize / 1000) Then Exit Function
    
ParsePrice = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAs32nds( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs32nds"
Dim failpoint As String
On Error GoTo Err

ParsePriceAs32nds = mPriceParser.ParsePriceAs32nds(pPriceString, pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAs32ndsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs32ndsAndFractions"
Dim failpoint As String
On Error GoTo Err

ParsePriceAs32ndsAndFractions = mPriceParser.ParsePriceAs32ndsAndFractions(pPriceString, pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAs64ths( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs64ths"
Dim failpoint As String
On Error GoTo Err

ParsePriceAs64ths = mPriceParser.ParsePriceAs64ths(pPriceString, pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAs64thsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "ParsePriceAs64ths"
Dim failpoint As String
On Error GoTo Err

ParsePriceAs64thsAndFractions = mPriceParser.ParsePriceAs64thsAndFractions(pPriceString, pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAsDecimals( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean

Const ProcName As String = "ParsePriceAsDecimals"
Dim failpoint As String
On Error GoTo Err

ParsePriceAsDecimals = mPriceParser.ParsePriceAsDecimals(pPriceString, pPrice)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub SaveToConfig( _
                ByVal config As ConfigurationSection)
                
Const ProcName As String = "SaveToConfig"
On Error GoTo Err

config.AddConfigurationSection ConfigSectionContractSpecifier
With config
    .SetSetting ConfigSettingContractSpecLocalSYmbol, mSpecifier.LocalSymbol
    .SetSetting ConfigSettingContractSpecSymbol, mSpecifier.Symbol
    .SetSetting ConfigSettingContractSpecExchange, mSpecifier.Exchange
    .SetSetting ConfigSettingContractSpecSecType, gSecTypeToString(mSpecifier.SecType)
    .SetSetting ConfigSettingContractSpecCurrency, mSpecifier.CurrencyCode
    .SetSetting ConfigSettingContractSpecExpiry, mSpecifier.Expiry
    .SetSetting ConfigSettingContractSpecStrikePrice, mSpecifier.Strike
    .SetSetting ConfigSettingContractSpecRight, gOptionRightToString(mSpecifier.Right)

    .SetSetting ConfigSettingDaysBeforeExpiryToSwitch, mDaysBeforeExpiryToSwitch
    .SetSetting ConfigSettingDescription, mDescription
    .SetSetting ConfigSettingExpiryDate, FormatTimestamp(mExpiryDate, TimestampDateOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingMultiplier, mMultiplier
    .SetSetting ConfigSettingSessionEndTime, FormatTimestamp(mSessionEndTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingSessionStartTime, FormatTimestamp(mSessionStartTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingSessionTickSize, mTickSize
    .SetSetting ConfigSettingSessionTimezone, mTimeZone.StandardName
End With


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function ToString() As String
Dim i As Long

Const ProcName As String = "ToString"
Dim failpoint As String
On Error GoTo Err

ToString = "Specifier=(" & mSpecifier.ToString & "); " & _
            "Description=" & mDescription & "; " & _
            "Expiry date=" & mExpiryDate & "; " & _
            "Tick size=" & mTickSize & "; " & _
            "Multiplier=" & mMultiplier & "; " & _
            "Session start=" & FormatDateTime(mSessionStartTime, vbShortTime) & "; " & _
            "Session end=" & FormatDateTime(mSessionEndTime, vbShortTime) & "; " & _
            "Timezone=" & mTimeZone.StandardName

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function ToXML() As String
Dim XMLdoc As DOMDocument30
Dim Contract As IXMLDOMElement
Dim Specifier As IXMLDOMElement
Dim ComboLegs As IXMLDOMElement
Dim ComboLeg As IXMLDOMElement
Dim comboLegObj As ComboLeg

Const ProcName As String = "ToXML"
Dim failpoint As String
On Error GoTo Err

Set XMLdoc = New DOMDocument30
Set Contract = XMLdoc.createElement("contract")
Set XMLdoc.documentElement = Contract
Contract.setAttribute "xmlns", "urn:tradewright.com:tradebuild"
Contract.setAttribute "minimumtick", mTickSize
Contract.setAttribute "multiplier", mMultiplier
Contract.setAttribute "sessionstarttime", Format(mSessionStartTime, "hh:mm:ss")
Contract.setAttribute "sessionendtime", Format(mSessionEndTime, "hh:mm:ss")
Contract.setAttribute "Description", mDescription
Contract.setAttribute "numberofdecimals", mNumberOfDecimals
Contract.setAttribute "timezonename", mTimeZone.StandardName

Set Specifier = XMLdoc.createElement("specifier")
Contract.appendChild Specifier
Specifier.setAttribute "symbol", mSpecifier.Symbol
Specifier.setAttribute "sectype", mSpecifier.SecType
Specifier.setAttribute "expiry", mSpecifier.Expiry
Specifier.setAttribute "exchange", mSpecifier.Exchange
Specifier.setAttribute "currencycode", mSpecifier.CurrencyCode
Specifier.setAttribute "localsymbol", mSpecifier.LocalSymbol
Specifier.setAttribute "right", mSpecifier.Right
Specifier.setAttribute "strike", mSpecifier.Strike

Set ComboLegs = XMLdoc.createElement("combolegs")
Specifier.appendChild ComboLegs
If Not mSpecifier.ComboLegs Is Nothing Then
    For Each comboLegObj In mSpecifier.ComboLegs
        Set ComboLeg = XMLdoc.createElement("comboleg")
        ComboLegs.appendChild ComboLeg
        
        Set Specifier = XMLdoc.createElement("specifier")
        ComboLeg.appendChild Specifier
        ComboLeg.setAttribute "isBuyLeg", comboLegObj.IsBuyLeg
        ComboLeg.setAttribute "Ratio", comboLegObj.Ratio
        Specifier.setAttribute "symbol", mSpecifier.Symbol
        Specifier.setAttribute "sectype", mSpecifier.SecType
        Specifier.setAttribute "expiry", mSpecifier.Expiry
        Specifier.setAttribute "exchange", mSpecifier.Exchange
        Specifier.setAttribute "currencycode", mSpecifier.CurrencyCode
        Specifier.setAttribute "localsymbol", mSpecifier.LocalSymbol
        Specifier.setAttribute "right", mSpecifier.Right
        Specifier.setAttribute "strike", mSpecifier.Strike
    Next
End If
ToXML = XMLdoc.xml

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

'@================================================================================
' Helper Functions
'@================================================================================

