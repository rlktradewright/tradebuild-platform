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

Implements IContract
Implements IJSONable
Implements IStringable

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

Private mSpecifier                                  As IContractSpecifier
Private mTickSize                                   As Double
Private mNumberOfDecimals                           As Long
Private mDescription                                As String
Private mSessionStartTime                           As Date
Private mSessionEndTime                             As Date
Private mFullSessionStartTime                       As Date
Private mFullSessionEndTime                         As Date
Private mExpiryDate                                 As Date
Private mDaysBeforeExpiryToSwitch                   As Long
Private mProviderIDs                                As Parameters
Private mTimezoneName                               As String

Private mProviderProperties                         As Parameters

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mDaysBeforeExpiryToSwitch = DefaultDaysBeforeExpiryToSwitch
mTickSize = DefaultTickSize
mTimezoneName = DefaultTimezoneName
End Sub

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

Private Property Get IContract_FullSessionEndTime() As Date
IContract_FullSessionEndTime = FullSessionEndTime
End Property

Private Property Get IContract_FullSessionStartTime() As Date
IContract_FullSessionStartTime = FullSessionStartTime
End Property

Private Property Get IContract_NumberOfDecimals() As Long
IContract_NumberOfDecimals = NumberOfDecimals
End Property

Private Property Get IContract_ProviderProperties() As Parameters
Set IContract_ProviderProperties = ProviderProperties
End Property

Private Property Get IContract_SessionEndTime() As Date
IContract_SessionEndTime = SessionEndTime
End Property

Private Property Get IContract_SessionStartTime() As Date
IContract_SessionStartTime = SessionStartTime
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
' IJSONable Interface Members
'@================================================================================

Private Function IJsonable_ToJSON() As String
IJsonable_ToJSON = ToJSON
End Function

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToString
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

Public Property Let FullSessionStartTime(ByVal Value As Date)
mFullSessionStartTime = TimeValue(Value)
End Property

Public Property Get FullSessionStartTime() As Date
FullSessionStartTime = mFullSessionStartTime
End Property

Public Property Let FullSessionEndTime(ByVal Value As Date)
mFullSessionEndTime = TimeValue(Value)
End Property

Public Property Get FullSessionEndTime() As Date
FullSessionEndTime = mFullSessionEndTime
End Property

Public Property Get Key() As String
Const ProcName As String = "Key"
On Error GoTo Err

Key = Specifier.Key

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Public Property Let ProviderProperties(ByVal Value As Parameters)
Set mProviderProperties = Value.Clone
End Property

Public Property Get ProviderProperties() As Parameters
If mProviderProperties Is Nothing Then Exit Property
Set ProviderProperties = mProviderProperties.Clone
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
Const ProcName As String = "TickSize"
On Error GoTo Err

Dim minTickString As String

mTickSize = Value

minTickString = Format(mTickSize, "0.##############")

mNumberOfDecimals = Len(minTickString) - 2

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickValue() As Double
TickValue = mTickSize * mSpecifier.Multiplier
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

Public Function FormatPrice(ByVal pPrice As Double) As String
Const ProcName As String = "FormatPrice"
On Error GoTo Err

' see http://www.cmegroup.com/trading/interest-rates/files/TreasuryFuturesPriceRoundingConventions_Mar_24_Final.pdf
' for details of price presentation, especially sections (2) and (7)

FormatPrice = gFormatPrice(pPrice, mSpecifier.SecType, mTickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub FromXML(ByVal contractXML As String)
Const ProcName As String = "FromXML"
On Error GoTo Err

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
Dim comboLegsObj As ComboLegs

Set XMLdoc = New DOMDocument30
XMLdoc.loadXML contractXML
Assert XMLdoc.parseError.errorCode = 0, "Contract XML cannot be parsed"

Set contractEl = XMLdoc.selectSingleNode("contract")
TickSize = contractEl.getAttribute("minimumtick")
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
If IsNull(contractEl.getAttribute("fullSessionstarttime")) Then
    mFullSessionStartTime = "00:00:00"
Else
    mFullSessionStartTime = contractEl.getAttribute("fullSessionstarttime")
End If
If IsNull(contractEl.getAttribute("fullSessionendtime")) Then
    mFullSessionEndTime = "00:00:00"
Else
    mFullSessionEndTime = contractEl.getAttribute("fullSessionendtime")
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
    Dim lMultiplier As Double
    If Not IsNull(contractEl.getAttribute("multiplier")) Then
        lMultiplier = contractEl.getAttribute("multiplier")
    Else
        lMultiplier = specifierEl.getAttribute("multiplier")
    End If
     
    .Initialise specifierEl.getAttribute("localsymbol"), _
                specifierEl.getAttribute("symbol"), _
                specifierEl.getAttribute("exchange"), _
                specifierEl.getAttribute("sectype"), _
                specifierEl.getAttribute("currencycode"), _
                specifierEl.getAttribute("expiry"), _
                lMultiplier, _
                specifierEl.getAttribute("strike"), _
                specifierEl.getAttribute("right")
    
    Set comboLegsEl = specifierEl.selectSingleNode("combolegs")
    Set nodeList = comboLegsEl.selectNodes("comboleg")
    If nodeList.length <> 0 Then .ComboLegs = New ComboLegs
    For Each comboLegEl In nodeList
        Set comboLegObj = New ComboLeg
        Set comboLegsObj = .ComboLegs
        comboLegsObj.Add comboLegObj
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

Public Sub SaveToConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "SaveToConfig"
On Error GoTo Err

gSaveContractToConfig Me, pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ToJSON() As String
Const ProcName As String = "ToJSON"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "Specifier"
jb.EncodeValue mSpecifier

jb.EncodeNameAndValue "DaysBeforeExpiryToSwitch", mDaysBeforeExpiryToSwitch, DefaultDaysBeforeExpiryToSwitch

jb.EncodeNameAndValue "Description", mDescription, ""

jb.EncodeNameAndValue "ExpiryDate", mExpiryDate, CDate(0#)

jb.EncodeNameAndValue "ProviderIDs", mProviderIDs, Nothing

jb.EncodeNameAndValue "SessionEndTime", mSessionEndTime, CDate(0#)

jb.EncodeNameAndValue "SessionStartTime", mSessionStartTime, CDate(0#)

jb.EncodeNameAndValue "FullSessionEndTime", mFullSessionEndTime, CDate(0#)

jb.EncodeNameAndValue "FullSessionStartTime", mFullSessionStartTime, CDate(0#)

jb.EncodeNameAndValue "TickSize", mTickSize, DefaultTickSize

jb.EncodeNameAndValue "TimezoneName", mTimezoneName, DefaultTimezoneName

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

ToString = gContractToString(Me)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ToXML() As String
Const ProcName As String = "ToXML"
On Error GoTo Err

ToXML = gContractToXML(Me)


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

