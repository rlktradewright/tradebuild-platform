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

Implements Comparable
Implements IContract
Implements JSONable
Implements Stringable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "Contract"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSpecifier As IContractSpecifier
Private mTickSize As Double
Private mMultiplier As Double
Private mNumberOfDecimals As Long
Private mDescription As String
Private mSessionStartTime As Date
Private mSessionEndTime As Date
Private mExpiryDate As Date
Private mDaysBeforeExpiryToSwitch As Long
Private mProviderIDs As Parameters
Private mTimezoneName As String

Private mSortKeys() As ContractSortKeyIds

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Comparable Interface Members
'@================================================================================

Private Function Comparable_CompareTo( _
                ByVal pObj As Comparable) As Long
Const ProcName As String = "Comparable_CompareTo"
On Error GoTo Err

Comparable_CompareTo = gContractsCompare(Me, pObj, mSortKeys)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' IContract Interface Members
'@================================================================================

Private Property Get IContract_DaysBeforeExpiryToSwitch() As Long
IContract_DaysBeforeExpiryToSwitch = DaysBeforeExpiryToSwitch
End Property

Private Property Get IContract_Description() As String
IContract_Description = Description
End Property

Private Property Get IContract_ExpiryDate() As Date
IContract_ExpiryDate = ExpiryDate
End Property

Private Property Get IContract_Multiplier() As Double
IContract_Multiplier = Multiplier
End Property

Private Property Get IContract_NumberOfDecimals() As Long
IContract_NumberOfDecimals = NumberOfDecimals
End Property

Private Property Get IContract_SessionEndTime() As Date
IContract_SessionEndTime = SessionEndTime
End Property

Private Property Get IContract_SessionStartTime() As Date
IContract_SessionStartTime = SessionStartTime
End Property

Private Property Let IContract_SortKeys(Value() As ContractSortKeyIds)
SortKeys = Value
End Property

Private Property Get IContract_Specifier() As IContractSpecifier
Set IContract_Specifier = Specifier
End Property

Private Property Get IContract_TickSize() As Double
IContract_TickSize = TickSize
End Property

Private Property Get IContract_TickValue() As Double
IContract_TickValue = TickValue
End Property

Private Property Get IContract_TimezoneName() As String
IContract_TimezoneName = TimezoneName
End Property

Private Function IContract_ToString() As String
IContract_ToString = ToString
End Function

'@================================================================================
' JSONable Interface Members
'@================================================================================

Private Function JSONable_ToJSON() As String
JSONable_ToJSON = ToJSON
End Function

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function Stringable_ToString() As String
Stringable_ToString = ToString
End Function

'@================================================================================
' Properties
'@================================================================================

Public Property Get DaysBeforeExpiryToSwitch() As Long
DaysBeforeExpiryToSwitch = mDaysBeforeExpiryToSwitch
End Property

Public Property Let DaysBeforeExpiryToSwitch(ByVal Value As Long)
mDaysBeforeExpiryToSwitch = Value
End Property

Public Property Let Description(ByVal Value As String)
mDescription = Value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Get ExpiryDate() As Date
ExpiryDate = mExpiryDate
End Property

Public Property Let ExpiryDate(ByVal Value As Date)
mExpiryDate = Value
End Property

Public Property Get Key() As String
Const ProcName As String = "Key"

On Error GoTo Err

Key = Specifier.Key

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Multiplier() As Double
Multiplier = mMultiplier
End Property

Public Property Let Multiplier(ByVal Value As Double)
mMultiplier = Value
End Property

Public Property Get NumberOfDecimals() As Long
NumberOfDecimals = mNumberOfDecimals
End Property

Public Property Get ProviderID(ByVal providerKey As String) As String
Const ProcName As String = "ProviderID"

On Error GoTo Err

ProviderID = mProviderIDs.GetParameterValue(providerKey, mSpecifier.LocalSymbol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ProviderIDs(ByVal Value As Parameters)
Set mProviderIDs = Value
End Property

Public Property Get ProviderIDs() As Parameters
Set ProviderIDs = mProviderIDs
End Property

Public Property Let SessionStartTime(ByVal Value As Date)
mSessionStartTime = TimeValue(Value)
End Property

Public Property Get SessionStartTime() As Date
SessionStartTime = mSessionStartTime
End Property

Public Property Let SessionEndTime(ByVal Value As Date)
mSessionEndTime = TimeValue(Value)
End Property

Public Property Get SessionEndTime() As Date
SessionEndTime = mSessionEndTime
End Property

Public Property Let SortKeys( _
                ByRef Value() As ContractSortKeyIds)
mSortKeys = Value
End Property
                
Public Property Get Specifier() As IContractSpecifier
Set Specifier = mSpecifier
End Property

Public Property Let Specifier(ByVal Value As IContractSpecifier)
Set mSpecifier = Value
End Property

Public Property Get TickSize() As Double
TickSize = mTickSize
End Property

Public Property Let TickSize(ByVal Value As Double)
Dim minTickString As String

Const ProcName As String = "TickSize"

On Error GoTo Err

mTickSize = Value

minTickString = Format(mTickSize, "0.##############")

mNumberOfDecimals = Len(minTickString) - 2

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get TickValue() As Double
TickValue = mTickSize * mMultiplier
End Property

Public Property Get TimezoneName() As String
TimezoneName = mTimezoneName
End Property

Public Property Let TimezoneName( _
                ByVal Value As String)
mTimezoneName = Value
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

Const ProcName As String = "FromXML"

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
    mTimezoneName = ""
Else
    mTimezoneName = contractEl.getAttribute("timezonename")
End If

Set specifierEl = contractEl.selectSingleNode("specifier")
Set specifierObj = New ContractSpecifier
Me.Specifier = specifierObj
With specifierObj
    .Initialise specifierEl.getAttribute("localsymbol"), _
                specifierEl.getAttribute("symbol"), _
                specifierEl.getAttribute("exchange"), _
                specifierEl.getAttribute("sectype"), _
                specifierEl.getAttribute("currencycode"), _
                specifierEl.getAttribute("expiry"), _
                specifierEl.getAttribute("strike"), _
                specifierEl.getAttribute("right")
    
    Set comboLegsEl = specifierEl.selectSingleNode("combolegs")
    Set nodeList = comboLegsEl.selectNodes("comboleg")
    If nodeList.length <> 0 Then .ComboLegs = New ComboLegs
    For Each comboLegEl In nodeList
        Set comboLegObj = New ComboLeg
        .ComboLegs.Add comboLegObj
        Set comboSpecifierEl = comboLegEl.selectSingleNode("specifier")
        Set comboSpecifierObj = New ContractSpecifier
        comboSpecifierObj.Initialise comboSpecifierEl.getAttribute("localsymbol"), _
                                    comboSpecifierEl.getAttribute("symbol"), _
                                    comboSpecifierEl.getAttribute("exchange"), _
                                    comboSpecifierEl.getAttribute("sectype"), _
                                    comboSpecifierEl.getAttribute("currencycode"), _
                                    comboSpecifierEl.getAttribute("expiry"), _
                                    comboSpecifierEl.getAttribute("strike"), _
                                    comboSpecifierEl.getAttribute("right")
        comboLegObj.Initialise comboSpecifierObj, _
                                CBool(comboLegEl.getAttribute("isbuyleg")), _
                                CLng(comboLegEl.getAttribute("Ratio"))
    Next

End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function IsValidPrice( _
                ByVal pPrice As Double, _
                ByVal pPrevValidPrice As Double) As Boolean
Const ProcName As String = "IsValidPrice"

On Error GoTo Err

IsValidPrice = gIsValidPrice(pPrice, pPrevValidPrice, mSpecifier.SecType, mTickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    .SetSetting ConfigSettingTickSize, mTickSize
    .SetSetting ConfigSettingTimezoneName, mTimezoneName
End With


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Const ProcName As String = "ToJSON"
On Error GoTo Err

Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "Specifier"
jb.EncodeValue mSpecifier

jb.EncodeName "DaysBeforeExpiryToSwitch"
jb.EncodeValue mDaysBeforeExpiryToSwitch

jb.EncodeName "Description"
jb.EncodeValue mDescription

jb.EncodeName "ExpiryDate"
jb.EncodeValue mExpiryDate

jb.EncodeName "Multiplier"
jb.EncodeValue mMultiplier

jb.EncodeName "ProviderIDs"
jb.EncodeValue mProviderIDs

jb.EncodeName "SessionEndTime"
jb.EncodeValue mSessionEndTime

jb.EncodeName "SessionStartTime"
jb.EncodeValue mSessionStartTime

jb.EncodeName "TickSize"
jb.EncodeValue mTickSize

jb.EncodeName "TimezoneName"
jb.EncodeValue mTimezoneName

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ToString() As String
Dim i As Long

Const ProcName As String = "ToString"

On Error GoTo Err

ToString = gContractToString(Me)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ToXML() As String
Dim XMLdoc As DOMDocument30
Dim Contract As IXMLDOMElement
Dim Specifier As IXMLDOMElement
Dim ComboLegs As IXMLDOMElement
Dim ComboLeg As IXMLDOMElement
Dim comboLegObj As ComboLeg

Const ProcName As String = "ToXML"

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
Contract.setAttribute "timezonename", mTimezoneName

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
gHandleUnexpectedError ProcName, ModuleName

End Function

'@================================================================================
' Helper Functions
'@================================================================================

