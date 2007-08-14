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

Private mSpecifier As ContractSpecifier
Private mTickSize As Double
Private mMultiplier As Long
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
Dim Contract As IXMLDOMElement
Dim specifier As IXMLDOMElement
Dim specifierObj As ContractSpecifier
Dim OrderTypes As IXMLDOMElement
Dim orderType As IXMLDOMElement
Dim exchanges As IXMLDOMElement
Dim exchange As IXMLDOMElement
Dim ComboLegs As IXMLDOMElement
Dim ComboLeg As IXMLDOMElement
Dim comboLegObj As ComboLeg
Dim nodeList As IXMLDOMNodeList
Dim comboSpecifier As IXMLDOMElement
Dim comboSpecifierObj As ContractSpecifier
Dim i As Long

Set XMLdoc = New DOMDocument30
XMLdoc.loadXML contractXML
If XMLdoc.parseError.errorCode <> 0 Then
    Err.Raise ErrorCodes.ErrRuntimeException, _
            "TradeBuild26.Contract::fromXML", _
            "Contract XML cannot be parsed"
End If

Set Contract = XMLdoc.selectSingleNode("contract")
TickSize = Contract.getAttribute("minimumtick")
mMultiplier = Contract.getAttribute("multiplier")
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
If IsNull(Contract.getAttribute("timezonename")) Then
    Set mTimeZone = GetTimeZone()
Else
    Set mTimeZone = GetTimeZone(CStr(Contract.getAttribute("timezonename")))
End If

Set specifier = Contract.selectSingleNode("specifier")
Set specifierObj = New ContractSpecifier
Me.specifier = specifierObj
With specifierObj
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
        .ComboLegs.Add ComboLeg
        Set comboSpecifier = ComboLeg.selectSingleNode("specifier")
        Set comboSpecifierObj = New ContractSpecifier
        With comboSpecifierObj
            .currencyCode = comboSpecifier.getAttribute("currencycode")
            .exchange = comboSpecifier.getAttribute("exchange")
            .expiry = comboSpecifier.getAttribute("expiry")
            .localSymbol = comboSpecifier.getAttribute("localsymbol")
            .right = comboSpecifier.getAttribute("right")
            .sectype = comboSpecifier.getAttribute("sectype")
            .strike = comboSpecifier.getAttribute("strike")
            .symbol = comboSpecifier.getAttribute("symbol")
        End With
        comboLegObj.initialise comboSpecifierObj, _
                                CBool(ComboLeg.getAttribute("isbuyleg")), _
                                CLng(ComboLeg.getAttribute("ratio"))
    Next

End With

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

ToXML = XMLdoc.xml

End Function

'@================================================================================
' Helper Functions
'@================================================================================






