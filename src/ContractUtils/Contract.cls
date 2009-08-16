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

Private Const ThirtySecondsSeparator As String = "'"

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
Private mPriceFormatString As String
Private mDescription As String
Private mSessionStartTime As Date
Private mSessionEndTime As Date
Private mExpiryDate As Date
Private mDaysBeforeExpiryToSwitch As Long
Private mProviderIDs As Parameters
Private mTimeZone As timezone

Private mSortKeys() As ContractSortKeyIds

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Comparable Interface Members
'@================================================================================

Private Function Comparable_compareTo( _
                ByVal obj As TWUtilities30.Comparable) As Long
Dim i As Long
Dim other As Contract
Dim otherSpec As ContractSpecifier

Set other = obj
Set otherSpec = other.specifier

For i = 0 To UBound(mSortKeys)
    Select Case mSortKeys(i)
    Case ContractSortKeyNone
        Exit Function
    Case ContractSortKeyLocalSymbol
        Comparable_compareTo = StrComp(mSpecifier.localSymbol, otherSpec.localSymbol, vbTextCompare)
    Case ContractSortKeySymbol
        Comparable_compareTo = StrComp(mSpecifier.symbol, otherSpec.symbol, vbTextCompare)
    Case ContractSortKeySecType
        Comparable_compareTo = StrComp(gSecTypeToShortString(mSpecifier.sectype), gSecTypeToShortString(otherSpec.sectype), vbTextCompare)
    Case ContractSortKeyExchange
        Comparable_compareTo = StrComp(mSpecifier.exchange, otherSpec.exchange, vbTextCompare)
    Case ContractSortKeyExpiry
        Comparable_compareTo = StrComp(mSpecifier.expiry, otherSpec.expiry, vbTextCompare)
    Case ContractSortKeyCurrency
        Comparable_compareTo = StrComp(mSpecifier.currencyCode, otherSpec.currencyCode, vbTextCompare)
    Case ContractSortKeyRight
        Comparable_compareTo = StrComp(gOptionRightToString(mSpecifier.right), gOptionRightToString(otherSpec.right), vbTextCompare)
    Case ContractSortKeyStrike
        Comparable_compareTo = Sgn(mSpecifier.strike - otherSpec.strike)
    End Select
    If Comparable_compareTo <> 0 Then Exit Function
Next
End Function

'@================================================================================
' Properties
'@================================================================================

Public Property Get daysBeforeExpiryToSwitch() As Long
daysBeforeExpiryToSwitch = mDaysBeforeExpiryToSwitch
End Property

Friend Property Let daysBeforeExpiryToSwitch(ByVal value As Long)
mDaysBeforeExpiryToSwitch = value
End Property

Friend Property Let description(ByVal value As String)
mDescription = value
End Property

Public Property Get description() As String
description = mDescription
End Property

Public Property Get expiryDate() As Date
expiryDate = mExpiryDate
End Property

Friend Property Let expiryDate(ByVal value As Date)
mExpiryDate = value
End Property

Public Property Get key() As String
key = specifier.key
End Property

Public Property Get multiplier() As Double
multiplier = mMultiplier
End Property

Friend Property Let multiplier(ByVal value As Double)
mMultiplier = value
End Property

Public Property Get numberOfDecimals() As Long
numberOfDecimals = mNumberOfDecimals
End Property

Public Property Get priceFormatString() As String
priceFormatString = mPriceFormatString
End Property

Public Property Get providerID(ByVal providerKey As String) As String
providerID = mProviderIDs.getParameterValue(providerKey, mSpecifier.localSymbol)
End Property

Friend Property Let providerIDs(ByRef value As Parameters)
Set mProviderIDs = value
End Property

Friend Property Get providerIDs() As Parameters
Set providerIDs = mProviderIDs
End Property

Friend Property Let sessionStartTime(ByVal value As Date)
mSessionStartTime = TimeValue(value)
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = mSessionStartTime
End Property

Friend Property Let sessionEndTime(ByVal value As Date)
mSessionEndTime = TimeValue(value)
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = mSessionEndTime
End Property

Friend Property Let SortKeys( _
                ByRef value() As ContractSortKeyIds)
mSortKeys = value
End Property
                
Public Property Get specifier() As ContractSpecifier
Set specifier = mSpecifier
End Property

Friend Property Let specifier(ByVal value As ContractSpecifier)
Set mSpecifier = value
End Property

Public Property Get TickSize() As Double
TickSize = mTickSize
End Property

Friend Property Let TickSize(ByVal value As Double)
Dim minTickString As String

mTickSize = value

minTickString = Format(mTickSize, "0.########")

mNumberOfDecimals = Len(minTickString) - 2
If mNumberOfDecimals = 0 Then
    mPriceFormatString = "0"
Else
    mPriceFormatString = "0." & String(mNumberOfDecimals, "0")
End If

End Property

Public Property Get tickValue() As Double
tickValue = mTickSize * mMultiplier
End Property

Public Property Get timezone() As timezone
Set timezone = mTimeZone
End Property

Friend Property Let timezone( _
                ByVal value As timezone)
Set mTimeZone = value
End Property

'@================================================================================
' Methods
'@================================================================================

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

Set XMLdoc = New DOMDocument30
XMLdoc.loadXML contractXML
If XMLdoc.parseError.errorCode <> 0 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            "TradeBuild26.Contract::fromXML", _
            "Contract XML cannot be parsed"
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
Me.specifier = specifierObj
With specifierObj
    .currencyCode = specifierEl.getAttribute("currencycode")
    .exchange = specifierEl.getAttribute("exchange")
    .expiry = specifierEl.getAttribute("expiry")
    .localSymbol = specifierEl.getAttribute("localsymbol")
    .right = specifierEl.getAttribute("right")
    .sectype = specifierEl.getAttribute("sectype")
    .strike = specifierEl.getAttribute("strike")
    .symbol = specifierEl.getAttribute("symbol")

    Set comboLegsEl = specifierEl.selectSingleNode("combolegs")
    Set nodeList = comboLegsEl.selectNodes("comboleg")
    If nodeList.length <> 0 Then .ComboLegs = New ComboLegs
    For Each comboLegEl In nodeList
        Set comboLegObj = New ComboLeg
        .ComboLegs.add comboLegObj
        Set comboSpecifierEl = comboLegEl.selectSingleNode("specifier")
        Set comboSpecifierObj = New ContractSpecifier
        With comboSpecifierObj
            .currencyCode = comboSpecifierEl.getAttribute("currencycode")
            .exchange = comboSpecifierEl.getAttribute("exchange")
            .expiry = comboSpecifierEl.getAttribute("expiry")
            .localSymbol = comboSpecifierEl.getAttribute("localsymbol")
            .right = comboSpecifierEl.getAttribute("right")
            .sectype = comboSpecifierEl.getAttribute("sectype")
            .strike = comboSpecifierEl.getAttribute("strike")
            .symbol = comboSpecifierEl.getAttribute("symbol")
        End With
        comboLegObj.initialise comboSpecifierObj, _
                                CBool(comboLegEl.getAttribute("isbuyleg")), _
                                CLng(comboLegEl.getAttribute("ratio"))
    Next

End With

End Sub

Public Function IsValidPrice( _
                ByVal price As Double, _
                ByVal prevValidPrice As Double) As Boolean
If prevValidPrice = 0 Or prevValidPrice = MaxDouble Then
    If Abs(price) / mTickSize > &H1FFFFF Then Exit Function ' note that Z index has ticksize 0.01
                                                            ' so we need to allow plenty of room
                                                            ' &H1FFFFF = 2097151
    If mSpecifier.sectype = SecTypeIndex Then
        ' don't do this check for indexes because some of them, such as TICK-NYSE, can have both
        ' positive, zero and negative values
    Else
        If price <= 0 Then Exit Function
        If price / mTickSize < 1 Then Exit Function ' catch occasional very small prices from IB
    End If
Else
    If Abs(prevValidPrice - price) / mTickSize > 32767 Then Exit Function
    If mSpecifier.sectype = SecTypeIndex Then
        ' don't do this check for indexes because some of them, such as TICK-NYSE, can have both
        ' positive and negative values - moreover the value can change dramatically from one
        ' tick to the next
    Else
        If price <= 0 Then Exit Function
        If price / mTickSize < 1 Then Exit Function ' catch occasional very small prices from IB
        If price < (0.5 * prevValidPrice) Or price > (1.5 * prevValidPrice) Then Exit Function
    End If
    
End If

IsValidPrice = True
End Function
                


Public Function parsePrice( _
                ByVal priceString As String, _
                ByRef price As Double) As Boolean
Static lRegExp As RegExp
Dim sepPosition As Long

If mTickSize = 0.03125 Then
    If lRegExp Is Nothing Then
        Set lRegExp = New RegExp
        lRegExp.Pattern = "^ *\d+ *($|" & ThirtySecondsSeparator & " *([0-9]|[0-2][0-9]|30|31) *$)"
    End If
    If Not lRegExp.Test(priceString) Then Exit Function
           
    sepPosition = InStr(1, priceString, ThirtySecondsSeparator)
    If sepPosition = 0 Then
        price = CDbl(priceString)
    Else
        price = CDbl(Left$(priceString, sepPosition - 1)) + _
                CDbl(right$(priceString, Len(priceString) - sepPosition)) / 32
    End If
Else
    If Not IsNumeric(priceString) Then Exit Function
    price = CDbl(priceString)
End If

If price <= 0 Then Exit Function

If Abs(price - mTickSize * Int((price + mTickSize / 1000) / mTickSize)) > (mTickSize / 1000) Then Exit Function
    
parsePrice = True
End Function

Public Function ToString() As String
Dim i As Long

ToString = "Specifier=(" & mSpecifier.ToString & "); " & _
            "Description=" & mDescription & "; " & _
            "Expiry date=" & mExpiryDate & "; " & _
            "Tick size=" & mTickSize & "; " & _
            "Multiplier=" & mMultiplier & "; " & _
            "Session start=" & FormatDateTime(mSessionStartTime, vbShortTime) & "; " & _
            "Session end=" & FormatDateTime(mSessionEndTime, vbShortTime) & "; " & _
            "Timezone=" & mTimeZone.standardName
End Function

Public Function ToXML() As String
Dim XMLdoc As DOMDocument30
Dim Contract As IXMLDOMElement
Dim specifier As IXMLDOMElement
Dim ComboLegs As IXMLDOMElement
Dim ComboLeg As IXMLDOMElement
Dim comboLegObj As ComboLeg

Set XMLdoc = New DOMDocument30
Set Contract = XMLdoc.createElement("contract")
Set XMLdoc.documentElement = Contract
Contract.setAttribute "xmlns", "urn:tradewright.com:tradebuild"
Contract.setAttribute "minimumtick", mTickSize
Contract.setAttribute "multiplier", mMultiplier
Contract.setAttribute "sessionstarttime", Format(mSessionStartTime, "hh:mm:ss")
Contract.setAttribute "sessionendtime", Format(mSessionEndTime, "hh:mm:ss")
Contract.setAttribute "description", mDescription
Contract.setAttribute "numberofdecimals", mNumberOfDecimals
Contract.setAttribute "timezonename", mTimeZone.standardName

Set specifier = XMLdoc.createElement("specifier")
Contract.appendChild specifier
specifier.setAttribute "symbol", mSpecifier.symbol
specifier.setAttribute "sectype", mSpecifier.sectype
specifier.setAttribute "expiry", mSpecifier.expiry
specifier.setAttribute "exchange", mSpecifier.exchange
specifier.setAttribute "currencycode", mSpecifier.currencyCode
specifier.setAttribute "localsymbol", mSpecifier.localSymbol
specifier.setAttribute "right", mSpecifier.right
specifier.setAttribute "strike", mSpecifier.strike

Set ComboLegs = XMLdoc.createElement("combolegs")
specifier.appendChild ComboLegs
If Not mSpecifier.ComboLegs Is Nothing Then
    For Each comboLegObj In mSpecifier.ComboLegs
        Set ComboLeg = XMLdoc.createElement("comboleg")
        ComboLegs.appendChild ComboLeg
        
        Set specifier = XMLdoc.createElement("specifier")
        ComboLeg.appendChild specifier
        ComboLeg.setAttribute "isBuyLeg", comboLegObj.isBuyLeg
        ComboLeg.setAttribute "ratio", comboLegObj.ratio
        specifier.setAttribute "symbol", mSpecifier.symbol
        specifier.setAttribute "sectype", mSpecifier.sectype
        specifier.setAttribute "expiry", mSpecifier.expiry
        specifier.setAttribute "exchange", mSpecifier.exchange
        specifier.setAttribute "currencycode", mSpecifier.currencyCode
        specifier.setAttribute "localsymbol", mSpecifier.localSymbol
        specifier.setAttribute "right", mSpecifier.right
        specifier.setAttribute "strike", mSpecifier.strike
    Next
End If
ToXML = XMLdoc.xml

End Function

'@================================================================================
' Helper Functions
'@================================================================================






