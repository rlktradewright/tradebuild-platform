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

Implements TradeBuildSP.IContract

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
Private mMarketName As String
Private mTradingClass As String
Private mContractID As Long
Private mTickSize As Double
Private mMultiplier As Long
Private mValidExchanges() As String
Private mNumberOfDecimals As Long
Private mPriceFormatString As String
Private mDescription As String
Private mSessionStartTime As Date
Private mSessionEndTime As Date
Private mExpiryDate As Date
Private mDaysBeforeExpiryToSwitch As Long
Private mProviderIDs() As TradeBuildSP.DictionaryEntry

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_ReadProperties(PropBag As PropertyBag)
Dim i As Long
Set mSpecifier = PropBag.ReadProperty("specifier")
mMarketName = PropBag.ReadProperty("marketname")
mTradingClass = PropBag.ReadProperty("tradingclass")
mContractID = PropBag.ReadProperty("conid")
mTickSize = PropBag.ReadProperty("mintick")
mMultiplier = PropBag.ReadProperty("multiplier")
ReDim mOrderTypes(CLng(PropBag.ReadProperty("uboundordertypes"))) As TradeBuild25.OrderTypes
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
PropBag.WriteProperty "mintick", mTickSize
PropBag.WriteProperty "multiplier", mMultiplier
PropBag.WriteProperty "uboundordertypes", 0
'PropBag.WriteProperty "uboundordertypes", UBound(mOrderTypes)
'For i = 0 To UBound(mOrderTypes)
'    PropBag.WriteProperty "ordertype" & i, mOrderTypes(i)
'Next
PropBag.WriteProperty "uboundvalidexchanges", UBound(mValidExchanges)
For i = 0 To UBound(mValidExchanges)
    PropBag.WriteProperty "validexchange" & i, mValidExchanges(i)
Next
End Sub

'@================================================================================
' IContract Interface Members
'@================================================================================

Private Function IContract_Clone() As TradeBuildSP.IContract
Set IContract_Clone = clone
End Function

Private Property Let IContract_ContractID(ByVal RHS As Long)
contractID = RHS
End Property

Private Property Get IContract_ContractID() As Long
IContract_ContractID = contractID
End Property

Private Property Let IContract_DaysBeforeExpiryToSwitch(ByVal RHS As Long)
daysBeforeExpiryToSwitch = RHS
End Property

Private Property Get IContract_DaysBeforeExpiryToSwitch() As Long
IContract_DaysBeforeExpiryToSwitch = daysBeforeExpiryToSwitch
End Property

Private Property Let IContract_Description(ByVal RHS As String)
Description = RHS
End Property

Private Property Get IContract_Description() As String
IContract_Description = Description
End Property

Private Property Let IContract_ExpiryDate(ByVal RHS As Date)
ExpiryDate = RHS
End Property

Private Property Get IContract_ExpiryDate() As Date
IContract_ExpiryDate = ExpiryDate
End Property

Private Sub IContract_FromXML( _
                ByVal contractXML As String)
FromXML contractXML
End Sub

Private Property Get IContract_Key() As String
IContract_Key = key
End Property

Private Property Let IContract_MarketName( _
                ByVal RHS As String)
marketName = RHS
End Property

Private Property Get IContract_MarketName() As String
IContract_MarketName = marketName
End Property

Private Property Let IContract_MinimumTick(ByVal RHS As Double)
TickSize = RHS
End Property

Private Property Get IContract_MinimumTick() As Double
IContract_MinimumTick = TickSize
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

Private Property Get IContract_ProviderID( _
                ByVal providerKey As String) As String
Dim i As Long
On Error GoTo err
For i = 0 To UBound(mProviderIDs)
    If mProviderIDs(i).key = providerKey Then
        IContract_ProviderID = mProviderIDs(i).value
        Exit Property
    End If
Next
IContract_ProviderID = mSpecifier.localSymbol
Exit Property

err:
IContract_ProviderID = mSpecifier.localSymbol
End Property

Private Property Let IContract_ProviderIDs( _
                ByRef RHS() As TradeBuildSP.DictionaryEntry)
mProviderIDs = RHS
End Property

Private Property Let IContract_SessionEndTime(ByVal RHS As Date)
sessionEndTime = TimeValue(RHS)
End Property

Private Property Get IContract_SessionEndTime() As Date
IContract_SessionEndTime = sessionEndTime
End Property

Private Property Let IContract_SessionStartTime( _
                ByVal RHS As Date)
sessionStartTime = TimeValue(RHS)
End Property

Private Property Get IContract_SessionStartTime() As Date
IContract_SessionStartTime = sessionStartTime
End Property

Private Property Let IContract_Specifier( _
                ByVal RHS As TradeBuildSP.IContractSpecifier)
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

Private Property Let IContract_TradingClass( _
                ByVal RHS As String)
tradingClass = RHS
End Property

Private Property Get IContract_TradingClass() As String
IContract_TradingClass = tradingClass
End Property

Private Property Let IContract_ValidExchanges( _
                RHS() As String)
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

'@================================================================================
' Properties
'@================================================================================

Public Property Get contractID() As Long
contractID = mContractID
End Property

Friend Property Let contractID(ByVal value As Long)
mContractID = value
End Property

Public Property Get daysBeforeExpiryToSwitch() As Long
daysBeforeExpiryToSwitch = mDaysBeforeExpiryToSwitch
End Property

Friend Property Let daysBeforeExpiryToSwitch(ByVal value As Long)
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

Public Property Get key() As String
key = specifier.key
End Property

Public Property Get marketName() As String
marketName = mMarketName
End Property

Friend Property Let marketName(ByVal value As String)
mMarketName = value
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

Public Property Get priceFormatString() As String
priceFormatString = mPriceFormatString
End Property

Friend Property Get providerID(ByVal providerKey As String) As String
Dim i As Long
On Error GoTo err
For i = 0 To UBound(mProviderIDs)
    If mProviderIDs(i).key = providerKey Then
        providerID = mProviderIDs(i).value
        Exit Property
    End If
Next
providerID = mSpecifier.localSymbol
Exit Property

err:
providerID = mSpecifier.localSymbol
End Property

Friend Property Let ProviderIDs(ByRef value() As TradeBuildSP.DictionaryEntry)
mProviderIDs = value
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

'@================================================================================
' Methods
'@================================================================================

Friend Function clone() As Contract

Set clone = New Contract

With clone
    .contractID = mContractID
    .daysBeforeExpiryToSwitch = mDaysBeforeExpiryToSwitch
    .Description = mDescription
    .ExpiryDate = mExpiryDate
    .marketName = mMarketName
    .TickSize = mTickSize
    .multiplier = mMultiplier
    '.OrderTypes = mOrderTypes
    .ProviderIDs = mProviderIDs
    .sessionEndTime = mSessionEndTime
    .sessionStartTime = sessionStartTime
    .specifier = mSpecifier.clone
    .tradingClass = mTradingClass
    .validExchanges = mValidExchanges
    .sessionEndTime = sessionEndTime
    .sessionStartTime = sessionStartTime
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
    err.Raise ErrorCodes.ErrRuntimeException, _
            "TradeBuild25.Contract::fromXML", _
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
    .locked = True
    .currencyCode = specifier.getAttribute("currencycode")
    .exchange = specifier.getAttribute("exchange")
    .expiry = specifier.getAttribute("expiry")
    .localSymbol = specifier.getAttribute("localsymbol")
    .right = specifier.getAttribute("right")
    .sectype = specifier.getAttribute("sectype")
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

Set OrderTypes = Contract.selectSingleNode("ordertypes")
Set nodeList = OrderTypes.selectNodes("ordertype")
ReDim mOrderTypes(nodeList.length) As TradeBuild25.OrderTypes
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
                CDbl(right$(priceString, Len(priceString) - sepPosition)) / 32
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




