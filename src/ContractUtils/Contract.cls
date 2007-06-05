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

Private mSpecifier As contractSpecifier
'Private mMarketName As String
'Private mTradingClass As String
'Private mContractID As Long
Private mTickSize As Double
Private mMultiplier As Long
'Private mValidExchanges() As String
Private mNumberOfDecimals As Long
Private mPriceFormatString As String
Private mDescription As String
Private mSessionStartTime As Date
Private mSessionEndTime As Date
Private mExpiryDate As Date
Private mDaysBeforeExpiryToSwitch As Long
Private mProviderIDs As Parameters
Private mTimeZone As timezone

'@================================================================================
' Class Event Handlers
'@================================================================================

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

Public Property Get multiplier() As Long
multiplier = mMultiplier
End Property

Friend Property Let multiplier(ByVal value As Long)
mMultiplier = value
End Property

Public Property Get numberOfDecimals() As Long
numberOfDecimals = mNumberOfDecimals
End Property

Public Property Get priceFormatString() As String
priceFormatString = mPriceFormatString
End Property

Friend Property Get providerID(ByVal providerKey As String) As String
providerID = mProviderIDs.getParameterValue(providerKey, mSpecifier.localsymbol)
End Property

Friend Property Let providerIDs(ByRef value As Parameters)
Set mProviderIDs = value
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

Public Property Get specifier() As contractSpecifier
Set specifier = mSpecifier
End Property

Friend Property Let specifier(ByVal value As contractSpecifier)
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

Friend Function clone() As Contract

Set clone = New Contract

With clone
    .contractID = mContractID
    .daysBeforeExpiryToSwitch = mDaysBeforeExpiryToSwitch
    .description = mDescription
    .expiryDate = mExpiryDate
    .marketName = mMarketName
    .TickSize = mTickSize
    .multiplier = mMultiplier
    '.OrderTypes = mOrderTypes
    .providerIDs = mProviderIDs
    .sessionEndTime = mSessionEndTime
    .sessionStartTime = sessionStartTime
    .specifier = mSpecifier.clone
    .tradingClass = mTradingClass
    .validExchanges = mValidExchanges
    .sessionEndTime = sessionEndTime
    .sessionStartTime = sessionStartTime
    .timezone = timezone
    '.NumberOfDecimals = mNumberOfDecimals  ' set automatically by tickSize
End With
End Function

Friend Sub FromXML(ByVal contractXML As String)
Dim XMLdoc As DOMDocument30
Dim Contract As IXMLDOMElement
Dim specifier As IXMLDOMElement
Dim specifierObj As contractSpecifier
Dim OrderTypes As IXMLDOMElement
Dim orderType As IXMLDOMElement
Dim exchanges As IXMLDOMElement
Dim exchange As IXMLDOMElement
Dim ComboLegs As IXMLDOMElement
Dim ComboLeg As IXMLDOMElement
Dim comboLegObj As ComboLeg
Dim nodeList As IXMLDOMNodeList
Dim i As Long

Set XMLdoc = New DOMDocument30
XMLdoc.loadXML contractXML
If XMLdoc.parseError.errorCode <> 0 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            "TradeBuild26.Contract::fromXML", _
            "Contract XML cannot be parsed"
End If

Set Contract = XMLdoc.selectSingleNode("contract")
mContractID = Contract.getAttribute("contractid")
mMarketName = Contract.getAttribute("marketname")
TickSize = Contract.getAttribute("minimumtick")
mMultiplier = Contract.getAttribute("multiplier")
mTradingClass = Contract.getAttribute("tradingclass")
If IsNull(Contract.getAttribute("sessionstarttime")) Then
    mSessionStartTime = "00:00:00"
Else
    mSessionStartTime = Contract.getAttribute("sessionstarttime")
End If
If IsNull(Contract.getAttribute("sessionendtime")) Then
    mSessionEndTime = "00:00:00"
Else
    mSessionEndTime = Contract.getAttribute("sessionendtime")
End If
If IsNull(Contract.getAttribute("description")) Then
    mDescription = ""
Else
    mDescription = Contract.getAttribute("description")
End If
If IsNull(Contract.getAttribute("numberofdecimals")) Then
    mNumberOfDecimals = 4
Else
    mNumberOfDecimals = Contract.getAttribute("numberofdecimals")
End If

Set specifier = Contract.selectSingleNode("specifier")
Set specifierObj = New contractSpecifier
Me.specifier = specifierObj
With specifierObj
    .Locked = True
    .currencyCode = specifier.getAttribute("currencycode")
    .exchange = specifier.getAttribute("exchange")
    .expiry = specifier.getAttribute("expiry")
    .localsymbol = specifier.getAttribute("localsymbol")
    .Right = specifier.getAttribute("right")
    .sectype = specifier.getAttribute("sectype")
    .strike = specifier.getAttribute("strike")
    .symbol = specifier.getAttribute("symbol")
    
    Set ComboLegs = specifier.selectSingleNode("combolegs")
    Set nodeList = ComboLegs.selectNodes("comboleg")
    For Each ComboLeg In nodeList
        Set comboLegObj = New ComboLeg
        .ComboLegs.Add ComboLeg
        comboLegObj.Action = ComboLeg.getAttribute("action")
        comboLegObj.contractID = ComboLeg.getAttribute("contractid")
        comboLegObj.exchange = ComboLeg.getAttribute("exchange")
        comboLegObj.openClose = ComboLeg.getAttribute("openclose")
        comboLegObj.ratio = ComboLeg.getAttribute("ratio")
    Next
    
End With

Set OrderTypes = Contract.selectSingleNode("ordertypes")
Set nodeList = OrderTypes.selectNodes("ordertype")
ReDim mOrderTypes(nodeList.length) As TradeBuild26.OrderTypes
i = 0
For Each orderType In nodeList
    mOrderTypes(i) = orderType.getAttribute("ordertype")
    i = i + 1
Next

Set exchanges = Contract.selectSingleNode("exchanges")
Set nodeList = exchanges.selectNodes("exchange")
ReDim mValidExchanges(nodeList.length) As String
i = 0
For Each exchange In nodeList
    mValidExchanges(i) = exchange.getAttribute("name")
    i = i + 1
Next

End Sub

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
                CDbl(Right$(priceString, Len(priceString) - sepPosition)) / 32
    End If
Else
    If Not IsNumeric(priceString) Then Exit Function
    price = CDbl(priceString)
End If

If price <= 0 Then Exit Function

If (price - mTickSize * Int(price / mTickSize)) <> 0 Then Exit Function
    
parsePrice = True
End Function

Public Function ToString() As String
Dim i As Long

ToString = mSpecifier.ToString

ToString = ToString & "exchanges:" & vbCrLf
For i = 0 To UBound(mValidExchanges)
    ToString = ToString & mValidExchanges(i) & IIf(i <> UBound(mValidExchanges), ",", "; ")
Next

End Function

Public Function ToXML() As String
Dim XMLdoc As DOMDocument30
Dim Contract As IXMLDOMElement
Dim specifier As IXMLDOMElement
Dim OrderTypes As IXMLDOMElement
'Dim orderType As IXMLDOMElement
Dim exchanges As IXMLDOMElement
Dim exchange As IXMLDOMElement
Dim ComboLegs As IXMLDOMElement
Dim ComboLeg As IXMLDOMElement
Dim comboLegObj As ComboLeg
Dim i As Long

Set XMLdoc = New DOMDocument30
Set Contract = XMLdoc.createElement("contract")
Set XMLdoc.documentElement = Contract
Contract.setAttribute "xmlns", "urn:tradewright.com:tradebuild"
Contract.setAttribute "contractid", mContractID
Contract.setAttribute "marketname", mMarketName
Contract.setAttribute "minimumtick", mTickSize
Contract.setAttribute "multiplier", mMultiplier
Contract.setAttribute "tradingclass", mTradingClass
Contract.setAttribute "sessionstarttime", Format(mSessionStartTime, "hh:mm:ss")
Contract.setAttribute "sessionendtime", Format(mSessionEndTime, "hh:mm:ss")
Contract.setAttribute "description", mDescription
Contract.setAttribute "numberofdecimals", mNumberOfDecimals

Set specifier = XMLdoc.createElement("specifier")
Contract.appendChild specifier
specifier.setAttribute "symbol", mSpecifier.symbol
specifier.setAttribute "sectype", mSpecifier.sectype
specifier.setAttribute "expiry", mSpecifier.expiry
specifier.setAttribute "exchange", mSpecifier.exchange
specifier.setAttribute "currencycode", mSpecifier.currencyCode
specifier.setAttribute "localsymbol", mSpecifier.localsymbol
specifier.setAttribute "right", mSpecifier.Right
specifier.setAttribute "strike", mSpecifier.strike

Set ComboLegs = XMLdoc.createElement("combolegs")
specifier.appendChild ComboLegs
For Each comboLegObj In mSpecifier.ComboLegs
    Set ComboLeg = XMLdoc.createElement("comboleg")
    ComboLegs.appendChild ComboLeg
    ComboLeg.setAttribute "action", comboLegObj.Action
    ComboLeg.setAttribute "contractid", comboLegObj.contractID
    ComboLeg.setAttribute "exchange", comboLegObj.exchange
    ComboLeg.setAttribute "openclose", comboLegObj.openClose
    ComboLeg.setAttribute "ratio", comboLegObj.ratio
Next

Set OrderTypes = XMLdoc.createElement("ordertypes")
Contract.appendChild OrderTypes
'For i = 0 To UBound(mOrderTypes)
'    Set orderType = XMLdoc.createElement("ordertype")
'    OrderTypes.appendChild orderType
'    orderType.setAttribute "ordertype", mOrderTypes(i)
'Next

Set exchanges = XMLdoc.createElement("exchanges")
Contract.appendChild exchanges
For i = 0 To UBound(mValidExchanges)
    Set exchange = XMLdoc.createElement("exchange")
    exchanges.appendChild exchange
    exchange.setAttribute "name", mValidExchanges(i)
Next

ToXML = XMLdoc.xml
End Function

'@================================================================================
' Helper Functions
'@================================================================================






