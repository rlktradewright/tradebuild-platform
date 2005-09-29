VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IContract

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mSpecifier As contractSpecifier
Private mMarketName As String
Private mTradingClass As String
Private mContractID As Long
Private mMinimumTick As Double
Private mMultiplier As Long
Private mOrderTypes() As TradeBuild.OrderTypes
Private mValidExchanges() As String
Private mTimestamp As String
Private mNumberOfDecimals As Long
Private mDescription As String
Private mSessionStartTime As Date
Private mSessionEndTime As Date

Private mReferenceTime As Date
Private mCurrentSessionStartTime As Date
Private mCurrentSessionEndTime As Date

Private mMarketIndex As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mMarketIndex = -1
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
Dim i As Long
Set mSpecifier = PropBag.ReadProperty("specifier")
mMarketName = PropBag.ReadProperty("marketname")
mTradingClass = PropBag.ReadProperty("tradingclass")
mContractID = PropBag.ReadProperty("conid")
mMinimumTick = PropBag.ReadProperty("mintick")
mMultiplier = PropBag.ReadProperty("multiplier")
ReDim mOrderTypes(CLng(PropBag.ReadProperty("uboundordertypes"))) As TradeBuild.OrderTypes
For i = 0 To UBound(mOrderTypes)
    mOrderTypes(i) = PropBag.ReadProperty("ordertype" & i)
Next
ReDim mValidExchanges(CLng(PropBag.ReadProperty("uboundvalidexchanges"))) As String
For i = 0 To UBound(mValidExchanges)
    mValidExchanges(i) = PropBag.ReadProperty("validexchange" & i)
Next
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
Dim i As Long
PropBag.WriteProperty "specifier", mSpecifier
PropBag.WriteProperty "marketname", mMarketName
PropBag.WriteProperty "tradingclass", mTradingClass
PropBag.WriteProperty "conid", mContractID
PropBag.WriteProperty "mintick", mMinimumTick
PropBag.WriteProperty "multiplier", mMultiplier
PropBag.WriteProperty "uboundordertypes", UBound(mOrderTypes)
For i = 0 To UBound(mOrderTypes)
    PropBag.WriteProperty "ordertype" & i, mOrderTypes(i)
Next
PropBag.WriteProperty "uboundvalidexchanges", UBound(mValidExchanges)
For i = 0 To UBound(mValidExchanges)
    PropBag.WriteProperty "validexchange" & i, mValidExchanges(i)
Next
End Sub

'================================================================================
' IContract Interface Members
'================================================================================

Private Function IContract_Clone() As TradeBuildSP.IContract
Set IContract_Clone = clone
End Function

Private Property Let IContract_ContractID(ByVal RHS As Long)
contractID = RHS
End Property

Private Property Get IContract_ContractID() As Long
IContract_ContractID = contractID
End Property

Private Property Let IContract_Description(ByVal RHS As String)
Description = RHS
End Property

Private Property Get IContract_Description() As String
IContract_Description = Description
End Property

Private Sub IContract_FromXML(ByVal contractXML As String)
FromXML contractXML
End Sub

Private Sub IContract_GetSessionTimes(ByVal timestamp As Date, sessionStartTime As Date, sessionEndTime As Date)
GetSessionTimes timestamp, sessionStartTime, sessionEndTime
End Sub

Private Function IContract_IsTimeInSession(ByVal timestamp As Date) As Boolean
IContract_IsTimeInSession = isTimeInSession(timestamp)
End Function

Private Property Get IContract_Key() As String
IContract_Key = key
End Property

Private Property Let IContract_MarketName(ByVal RHS As String)
marketName = RHS
End Property

Private Property Get IContract_MarketName() As String
IContract_MarketName = marketName
End Property

Private Property Let IContract_MinimumTick(ByVal RHS As Double)
minimumTick = RHS
End Property

Private Property Get IContract_MinimumTick() As Double
IContract_MinimumTick = minimumTick
End Property

Private Property Let IContract_Multiplier(ByVal RHS As Long)
multiplier = RHS
End Property

Private Property Get IContract_Multiplier() As Long
IContract_Multiplier = multiplier
End Property

Private Property Get IContract_NumberOfDecimals() As Long
IContract_NumberOfDecimals = NumberOfDecimals
End Property

Private Property Let IContract_OrderTypes(RHS() As TradeBuildSP.OrderTypes)
Dim lOrderTypes() As TradeBuild.OrderTypes
Dim i As Long

ReDim lOrderTypes(UBound(RHS)) As TradeBuild.OrderTypes
For i = 0 To UBound(RHS)
    Select Case RHS(i)
    Case TradeBuildSP.OrderTypes.OrderTypeNone
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeNone
    Case TradeBuildSP.OrderTypes.OrderTypeMarket
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeMarket
    Case TradeBuildSP.OrderTypes.OrderTypeMarketClose
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeMarketClose
    Case TradeBuildSP.OrderTypes.OrderTypeLimit
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeLimit
    Case TradeBuildSP.OrderTypes.OrderTypeLimitClose
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeLimitClose
    Case TradeBuildSP.OrderTypes.OrderTypePegMarket
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypePegMarket
    Case TradeBuildSP.OrderTypes.OrderTypeStop
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeStop
    Case TradeBuildSP.OrderTypes.OrderTypeStopLimit
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeStopLimit
    Case TradeBuildSP.OrderTypes.OrderTypeTrail
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeTrail
    Case TradeBuildSP.OrderTypes.OrderTypeRelative
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeRelative
    Case TradeBuildSP.OrderTypes.OrderTypeVWAP
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeVWAP
    Case TradeBuildSP.OrderTypes.OrderTypeMarketToLimit
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeMarketToLimit
    Case TradeBuildSP.OrderTypes.OrderTypeQuote
        lOrderTypes(i) = TradeBuild.OrderTypes.OrderTypeQuote
    End Select
Next
OrderTypes = lOrderTypes
End Property

Private Property Get IContract_OrderTypes() As TradeBuildSP.OrderTypes()
Dim lOrderTypes() As TradeBuildSP.OrderTypes
Dim ot() As TradeBuild.OrderTypes
Dim i As Long

ot = OrderTypes
ReDim lOrderTypes(UBound(ot)) As TradeBuildSP.OrderTypes
For i = 0 To UBound(ot)
    Select Case ot(i)
    Case TradeBuild.OrderTypes.OrderTypeNone
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeNone
    Case TradeBuild.OrderTypes.OrderTypeMarket
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeMarket
    Case TradeBuild.OrderTypes.OrderTypeMarketClose
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeMarketClose
    Case TradeBuild.OrderTypes.OrderTypeLimit
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeLimit
    Case TradeBuild.OrderTypes.OrderTypeLimitClose
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeLimitClose
    Case TradeBuild.OrderTypes.OrderTypePegMarket
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypePegMarket
    Case TradeBuild.OrderTypes.OrderTypeStop
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeStop
    Case TradeBuild.OrderTypes.OrderTypeStopLimit
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeStopLimit
    Case TradeBuild.OrderTypes.OrderTypeTrail
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeTrail
    Case TradeBuild.OrderTypes.OrderTypeRelative
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeRelative
    Case TradeBuild.OrderTypes.OrderTypeVWAP
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeVWAP
    Case TradeBuild.OrderTypes.OrderTypeMarketToLimit
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeMarketToLimit
    Case TradeBuild.OrderTypes.OrderTypeQuote
        lOrderTypes(i) = TradeBuildSP.OrderTypes.OrderTypeQuote
    End Select
Next
IContract_OrderTypes = lOrderTypes
End Property

Private Property Let IContract_SessionEndTime(ByVal RHS As Date)
sessionEndTime = RHS
End Property

Private Property Get IContract_SessionEndTime() As Date
IContract_SessionEndTime = sessionEndTime
End Property

Private Property Let IContract_SessionStartTime(ByVal RHS As Date)
sessionStartTime = RHS
End Property

Private Property Get IContract_SessionStartTime() As Date
IContract_SessionStartTime = sessionStartTime
End Property

Private Property Let IContract_Specifier(ByVal RHS As TradeBuildSP.IContractSpecifier)
specifier = RHS
End Property

Private Property Get IContract_Specifier() As TradeBuildSP.IContractSpecifier
Set IContract_Specifier = specifier
End Property

Private Function IContract_ToString() As String
IContract_ToString = ToString
End Function

Private Function IContract_ToXML() As String
IContract_ToXML = ToXML
End Function

Private Property Let IContract_TradingClass(ByVal RHS As String)
tradingClass = RHS
End Property

Private Property Get IContract_TradingClass() As String
IContract_TradingClass = tradingClass
End Property

Private Property Let IContract_ValidExchanges(RHS() As String)
Dim exchanges() As String
Dim i As Long

ReDim exchanges(UBound(RHS)) As String

For i = 0 To UBound(RHS)
    exchanges(i) = RHS(i)
Next

validExchanges = exchanges
End Property

Private Property Get IContract_ValidExchanges() As String()
IContract_ValidExchanges = validExchanges
End Property

'================================================================================
' Properties
'================================================================================

Public Property Get contractID() As Long
contractID = mContractID
End Property

Friend Property Let contractID(ByVal value As Long)
mContractID = value
End Property

Public Property Get currentSessionEndTime() As Date
currentSessionEndTime = mCurrentSessionEndTime
End Property

Public Property Get currentSessionStartTime() As Date
currentSessionStartTime = mCurrentSessionStartTime
End Property

Friend Property Let Description(ByVal value As String)
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Get key() As String
key = specifier.key
End Property

Public Property Get marketName() As String
marketName = mMarketName
End Property

Friend Property Let marketName(ByVal value As String)
mMarketName = value
End Property

Public Property Get minimumTick() As Double
minimumTick = mMinimumTick
End Property

Friend Property Let minimumTick(ByVal value As Double)
Dim minTickString As String

mMinimumTick = value

minTickString = Format(mMinimumTick, "0.########")

mNumberOfDecimals = Len(minTickString) - 2
End Property

Public Property Get multiplier() As Long
multiplier = mMultiplier
End Property

Friend Property Let multiplier(ByVal value As Long)
mMultiplier = value
End Property

Public Property Get NumberOfDecimals() As Long
NumberOfDecimals = mNumberOfDecimals
End Property

Public Property Get OrderTypes() As TradeBuild.OrderTypes()
OrderTypes = mOrderTypes
End Property

Friend Property Let OrderTypes(ByRef value() As TradeBuild.OrderTypes)
mOrderTypes = value
End Property

Friend Property Let sessionStartTime(ByVal value As Date)
mSessionStartTime = value
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = mSessionStartTime
End Property

Friend Property Let sessionEndTime(ByVal value As Date)
mSessionEndTime = value
End Property

Public Property Get sessionEndTime() As Date
sessionEndTime = mSessionEndTime
End Property

Friend Property Get marketIndex() As Long
marketIndex = mMarketIndex
End Property

Friend Property Let marketIndex(ByVal value As Long)
mMarketIndex = value
End Property

Public Property Get specifier() As contractSpecifier
Set specifier = mSpecifier
End Property

Friend Property Let specifier(ByVal value As contractSpecifier)
Set mSpecifier = value
End Property

Friend Property Let timestamp(ByVal value As Date)
mTimestamp = value
End Property

Friend Property Get timestamp() As Date
timestamp = mTimestamp
End Property

Public Property Get tradingClass() As String
tradingClass = mTradingClass
End Property

Friend Property Let tradingClass(ByVal value As String)
mTradingClass = value
End Property

Public Property Get validExchanges() As String()
validExchanges = mValidExchanges
End Property

Friend Property Let validExchanges(ByRef value() As String)
mValidExchanges = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function BarStartTime( _
                            ByVal timestamp As Date, _
                            ByVal barLength As Long) As Date
Dim offset As Long
Dim theDate As Long
Dim theTime As Double
Dim theTimeMins As Long

offset = DatePart("n", mSessionStartTime)
theDate = Fix(CDbl(timestamp))
' NB: don't use TimeValue to get the time, as VB rounds it to
' the nearest second
theTime = CDbl(timestamp) - theDate
theTimeMins = Fix(theTime * 1440) ' minutes since midnight
BarStartTime = theDate + _
            (barLength * Fix((theTimeMins - offset) / barLength) + _
                offset) / 1440

End Function

Public Sub GetSessionTimes( _
                            ByVal timestamp As Date, _
                            ByRef sessionStartTime As Date, _
                            ByRef sessionEndTime As Date)
CalcSessionTimes timestamp, sessionStartTime, sessionEndTime
End Sub

Friend Function clone() As contract

Set clone = New contract

With clone
    .contractID = mContractID
    .marketName = mMarketName
    .minimumTick = mMinimumTick
    .multiplier = mMultiplier
    .OrderTypes = mOrderTypes
    .specifier = mSpecifier.clone
    .tradingClass = mTradingClass
    .validExchanges = mValidExchanges
    .Description = mDescription
    .sessionEndTime = mSessionEndTime
    .sessionStartTime = sessionStartTime
    '.NumberOfDecimals = mNumberOfDecimals  ' set automatically by minimumTick
End With
End Function

Friend Sub FromXML(ByVal contractXML As String)
Dim XMLdoc As DOMDocument30
Dim contract As IXMLDOMElement
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
    err.Raise ErrorCodes.ContractCannotBeParsed, _
            "TradeBuild.Contract::fromXML", _
            "Contract XML cannot be parsed"
End If

Set contract = XMLdoc.selectSingleNode("contract")
mContractID = contract.getAttribute("contractid")
mMarketName = contract.getAttribute("marketname")
minimumTick = contract.getAttribute("minimumtick")
mMultiplier = contract.getAttribute("multiplier")
mTradingClass = contract.getAttribute("tradingclass")
mSessionStartTime = contract.getAttribute("sessionstarttime")
mSessionEndTime = contract.getAttribute("sessionendtime")
mDescription = contract.getAttribute("description")
mNumberOfDecimals = contract.getAttribute("numberofdecimals")

Set specifier = contract.selectSingleNode("specifier")
Set specifierObj = New contractSpecifier
Me.specifier = specifierObj
With specifierObj
    .locked = True
    .currencyCode = specifier.getAttribute("currencycode")
    .exchange = specifier.getAttribute("exchange")
    .expiry = specifier.getAttribute("expiry")
    .localSymbol = specifier.getAttribute("localsymbol")
    .right = specifier.getAttribute("right")
    .secType = specifier.getAttribute("sectype")
    .strike = specifier.getAttribute("strike")
    .symbol = specifier.getAttribute("symbol")
    
    Set ComboLegs = specifier.selectSingleNode("combolegs")
    Set nodeList = ComboLegs.selectNodes("comboleg")
    For Each ComboLeg In nodeList
        Set comboLegObj = New ComboLeg
        .ComboLegs.add ComboLeg
        comboLegObj.action = ComboLeg.getAttribute("action")
        comboLegObj.contractID = ComboLeg.getAttribute("contractid")
        comboLegObj.exchange = ComboLeg.getAttribute("exchange")
        comboLegObj.openClose = ComboLeg.getAttribute("openclose")
        comboLegObj.ratio = ComboLeg.getAttribute("ratio")
    Next
    
End With

Set OrderTypes = contract.selectSingleNode("ordertypes")
Set nodeList = OrderTypes.selectNodes("ordertype")
ReDim mOrderTypes(nodeList.length) As TradeBuild.OrderTypes
i = 0
For Each orderType In nodeList
    mOrderTypes(i) = orderType.getAttribute("ordertype")
    i = i + 1
Next

Set exchanges = contract.selectSingleNode("exchanges")
Set nodeList = exchanges.selectNodes("exchange")
ReDim mValidExchanges(nodeList.length) As String
i = 0
For Each exchange In nodeList
    mValidExchanges(i) = exchange.getAttribute("name")
    i = i + 1
Next

End Sub

Public Function isTimeInSession(ByVal timestamp As Date) As Boolean

If timestamp >= mCurrentSessionStartTime And _
    timestamp < mCurrentSessionEndTime _
Then
    isTimeInSession = True
End If
End Function

Public Function newOrder() As Order
Dim lOrder As Order
Set lOrder = New Order
lOrder.id = gNextOrderID
gAllOrders.add lOrder, CStr(gNextOrderID)
gNextOrderID = gNextOrderID + 1
lOrder.contract = Me
lOrder.marketIndex = mMarketIndex
Set newOrder = lOrder.clone   ' return a copy of the order to the app
End Function

Friend Sub setSession( _
                ByVal timestamp As Date)
If timestamp >= mCurrentSessionEndTime Or _
    timestamp < mReferenceTime _
Then
    mReferenceTime = timestamp
    CalcSessionTimes timestamp, mCurrentSessionStartTime, mCurrentSessionEndTime
End If
End Sub

Public Function ToString() As String
Dim i As Long

ToString = mSpecifier.ToString
ToString = ToString & "ordertypes:" & vbCrLf

For i = 0 To UBound(mOrderTypes)
    Select Case mOrderTypes(i)
    Case OrderTypeMarket
        ToString = ToString & vbTab & "OrderTypeMarket" & vbCrLf
    Case OrderTypeMarketClose
        ToString = ToString & vbTab & "OrderTypeMarketClose" & vbCrLf
    Case OrderTypeLimit
        ToString = ToString & vbTab & "OrderTypeLimit" & vbCrLf
    Case OrderTypeLimitClose
        ToString = ToString & vbTab & "OrderTypeLimitClose" & vbCrLf
    Case OrderTypePegMarket
        ToString = ToString & vbTab & "OrderTypePegMarket" & vbCrLf
    Case OrderTypeStop
        ToString = ToString & vbTab & "OrderTypeStop" & vbCrLf
    Case OrderTypeStopLimit
        ToString = ToString & vbTab & "OrderTypeStopLimit" & vbCrLf
    Case OrderTypeTrail
        ToString = ToString & vbTab & "OrderTypeTrail" & vbCrLf
    Case OrderTypeRelative
        ToString = ToString & vbTab & "OrderTypeRelative" & vbCrLf
    Case OrderTypeVWAP
        ToString = ToString & vbTab & "OrderTypeVWAP" & vbCrLf
    Case OrderTypeMarketToLimit
        ToString = ToString & vbTab & "OrderTypeMarketToLimit" & vbCrLf
    Case OrderTypeQuote
        ToString = ToString & vbTab & "OrderTypeQuote" & vbCrLf
    End Select
Next

ToString = ToString & "exchanges:" & vbCrLf
For i = 0 To UBound(mValidExchanges)
    ToString = ToString & vbTab & mValidExchanges(i) & vbCrLf
Next

End Function

Public Function ToXML() As String
Dim XMLdoc As DOMDocument30
Dim contract As IXMLDOMElement
Dim specifier As IXMLDOMElement
Dim OrderTypes As IXMLDOMElement
Dim orderType As IXMLDOMElement
Dim exchanges As IXMLDOMElement
Dim exchange As IXMLDOMElement
Dim ComboLegs As IXMLDOMElement
Dim ComboLeg As IXMLDOMElement
Dim comboLegObj As ComboLeg
Dim i As Long

Set XMLdoc = New DOMDocument30
Set contract = XMLdoc.createElement("contract")
Set XMLdoc.documentElement = contract
contract.setAttribute "xmlns", "urn:tradewright.com:tradebuild"
contract.setAttribute "contractid", mContractID
contract.setAttribute "marketname", mMarketName
contract.setAttribute "minimumtick", mMinimumTick
contract.setAttribute "multiplier", mMultiplier
contract.setAttribute "tradingclass", mTradingClass
contract.setAttribute "sessionstarttime", Format(mSessionStartTime, "hh:mm:ss")
contract.setAttribute "sessionendtime", Format(mSessionEndTime, "hh:mm:ss")
contract.setAttribute "description", mDescription
contract.setAttribute "numberofdecimals", mNumberOfDecimals

Set specifier = XMLdoc.createElement("specifier")
contract.appendChild specifier
specifier.setAttribute "symbol", mSpecifier.symbol
specifier.setAttribute "sectype", mSpecifier.secType
specifier.setAttribute "expiry", mSpecifier.expiry
specifier.setAttribute "exchange", mSpecifier.exchange
specifier.setAttribute "currencycode", mSpecifier.currencyCode
specifier.setAttribute "localsymbol", mSpecifier.localSymbol
specifier.setAttribute "right", mSpecifier.right
specifier.setAttribute "strike", mSpecifier.strike

Set ComboLegs = XMLdoc.createElement("combolegs")
specifier.appendChild ComboLegs
For Each comboLegObj In mSpecifier.ComboLegs
    Set ComboLeg = XMLdoc.createElement("comboleg")
    ComboLegs.appendChild ComboLeg
    ComboLeg.setAttribute "action", comboLegObj.action
    ComboLeg.setAttribute "contractid", comboLegObj.contractID
    ComboLeg.setAttribute "exchange", comboLegObj.exchange
    ComboLeg.setAttribute "openclose", comboLegObj.openClose
    ComboLeg.setAttribute "ratio", comboLegObj.ratio
Next

Set OrderTypes = XMLdoc.createElement("ordertypes")
contract.appendChild OrderTypes
For i = 0 To UBound(mOrderTypes)
    Set orderType = XMLdoc.createElement("ordertype")
    OrderTypes.appendChild orderType
    orderType.setAttribute "ordertype", mOrderTypes(i)
Next

Set exchanges = XMLdoc.createElement("exchanges")
contract.appendChild exchanges
For i = 0 To UBound(mValidExchanges)
    Set exchange = XMLdoc.createElement("exchange")
    exchanges.appendChild exchange
    exchange.setAttribute "name", mValidExchanges(i)
Next

ToXML = XMLdoc.xml
End Function

'================================================================================
' Helper Functions
'================================================================================

Friend Sub CalcSessionTimes(ByVal timestamp As Date, _
                            ByRef sessionStartTime As Date, _
                            ByRef sessionEndTime As Date)
Dim referenceDate As Date
Dim referenceTime As Date

referenceDate = DateValue(timestamp)
referenceTime = TimeValue(timestamp)

If mSessionStartTime < mSessionEndTime Then
    ' session doesn't span midnight
    If referenceTime < mSessionEndTime Then
        sessionStartTime = referenceDate + mSessionStartTime
        sessionEndTime = referenceDate + mSessionEndTime
    Else
        sessionStartTime = referenceDate + 1 + mSessionStartTime
        sessionEndTime = referenceDate + 1 + mSessionEndTime
    End If
ElseIf mSessionStartTime > mSessionEndTime Then
    ' session spans midnight
    If referenceTime >= mSessionEndTime Then
        sessionStartTime = referenceDate + mSessionStartTime
        sessionEndTime = referenceDate + 1 + mSessionEndTime
    Else
        sessionStartTime = referenceDate - 1 + mSessionStartTime
        sessionEndTime = referenceDate + mSessionEndTime
    End If
Else
    ' this instrument trades 24hrs
    sessionStartTime = "01/01/1900"
    sessionEndTime = "31/12/2099"
End If

End Sub



