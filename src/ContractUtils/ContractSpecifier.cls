VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractSpecifier"
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

Implements IContractSpecifier
Implements IJSONable
Implements IStringable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "ContractSpecifier"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSymbol         As String
Private mSecType        As SecurityTypes
Private mExpiry         As String
Private mStrike         As Double
Private mRight          As OptionRights
Private mExchange       As String
Private mCurrency       As String
Private mLocalSymbol    As String
Private mMultiplier     As Double
Private mComboLegs      As IComboLegs
Private mTradingClass   As String

Private mProviderProperties                         As Parameters

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mMultiplier = DefaultMultiplier
End Sub

'@================================================================================
' IContractSpecifier Interface Members
'@================================================================================

Private Property Get IContractSpecifier_ComboLegs() As IComboLegs
Set IContractSpecifier_ComboLegs = ComboLegs
End Property

Private Property Get IContractSpecifier_CurrencyCode() As String
IContractSpecifier_CurrencyCode = CurrencyCode
End Property

Private Function IContractSpecifier_Equals(ByVal pContractSpec As IContractSpecifier) As Boolean
IContractSpecifier_Equals = Equals(pContractSpec)
End Function

Private Property Get IContractSpecifier_Exchange() As String
IContractSpecifier_Exchange = Exchange
End Property

Private Property Get IContractSpecifier_Expiry() As String
IContractSpecifier_Expiry = Expiry
End Property

Private Property Get IContractSpecifier_Key() As String
IContractSpecifier_Key = Key
End Property

Private Property Get IContractSpecifier_LocalSymbol() As String
IContractSpecifier_LocalSymbol = LocalSymbol
End Property

Private Property Get IContractSpecifier_Multiplier() As Double
IContractSpecifier_Multiplier = Multiplier
End Property

Private Property Get IContractSpecifier_ProviderProperties() As Parameters
Set IContractSpecifier_ProviderProperties = ProviderProperties
End Property

Private Property Get IContractSpecifier_Right() As OptionRights
IContractSpecifier_Right = Right
End Property

Private Property Get IContractSpecifier_SecType() As SecurityTypes
IContractSpecifier_SecType = SecType
End Property

Private Property Get IContractSpecifier_Strike() As Double
IContractSpecifier_Strike = Strike
End Property

Private Property Get IContractSpecifier_Symbol() As String
IContractSpecifier_Symbol = Symbol
End Property

Private Function IContractSpecifier_ToString() As String
IContractSpecifier_ToString = ToString
End Function

Private Property Get IContractSpecifier_TradingClass() As String
IContractSpecifier_TradingClass = TradingClass
End Property

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

Public Property Let ComboLegs(ByVal Value As IComboLegs)
Set mComboLegs = Value
End Property

Public Property Get ComboLegs() As IComboLegs
Set ComboLegs = mComboLegs
End Property

Public Property Let zCurrencyCode(ByVal Value As String)
Attribute zCurrencyCode.VB_MemberFlags = "40"
mCurrency = UCase$(Value)
End Property

Public Property Get CurrencyCode() As String
CurrencyCode = mCurrency
End Property

Public Property Let zExchange(ByVal Value As String)
Attribute zExchange.VB_MemberFlags = "40"
mExchange = UCase$(Value)
End Property

Public Property Get Exchange() As String
Exchange = mExchange
End Property

Public Property Let zExpiry(ByVal Value As String)
Attribute zExpiry.VB_MemberFlags = "40"
Const ProcName As String = "Expiry"
On Error GoTo Err

If Value = "" Then Exit Property
Dim lErrorMessage As String
AssertArgument GContractUtils.IsValidExpiry(Value, lErrorMessage), "Invalid Expiry"

mExpiry = Value

Exit Property

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Expiry() As String
Expiry = mExpiry
End Property

Public Property Get Key() As String
Const ProcName As String = "Key"
On Error GoTo Err

Static sKey As String
If sKey = "" Then sKey = GContractUtils.GetContractSpecKey(Me)
Key = sKey

Exit Property

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let zLocalSymbol(ByVal Value As String)
Attribute zLocalSymbol.VB_MemberFlags = "40"
mLocalSymbol = UCase$(Value)
End Property

Public Property Get LocalSymbol() As String
LocalSymbol = mLocalSymbol
End Property

Public Property Let zMultiplier(ByVal Value As Double)
Attribute zMultiplier.VB_MemberFlags = "40"
mMultiplier = Value
End Property

Public Property Let Multiplier(ByVal Value As Double)
mMultiplier = Value
End Property

Public Property Get Multiplier() As Double
Multiplier = mMultiplier
End Property

Public Property Let ProviderProperties(ByVal Value As Parameters)
If Value Is Nothing Then
    Set mProviderProperties = Nothing
    Exit Property
End If
Set mProviderProperties = Value.Clone
End Property

Public Property Get ProviderProperties() As Parameters
If mProviderProperties Is Nothing Then Exit Property
Set ProviderProperties = mProviderProperties.Clone
End Property

Public Property Let zRight(ByVal Value As OptionRights)
Attribute zRight.VB_MemberFlags = "40"
mRight = Value
End Property

Public Property Get Right() As OptionRights
Right = mRight
End Property

Public Property Let zSecType(ByVal Value As SecurityTypes)
Attribute zSecType.VB_MemberFlags = "40"
mSecType = Value
End Property

Public Property Get SecType() As SecurityTypes
SecType = mSecType
End Property

Public Property Let zStrike(ByVal Value As Double)
Attribute zStrike.VB_MemberFlags = "40"
mStrike = Value
End Property

Public Property Get Strike() As Double
Strike = mStrike
End Property

Public Property Let zSymbol(ByVal Value As String)
Attribute zSymbol.VB_MemberFlags = "40"
mSymbol = UCase$(Value)
End Property

Public Property Get Symbol() As String
Symbol = mSymbol
End Property

Public Property Let zTradingClass(ByVal Value As String)
Attribute zTradingClass.VB_MemberFlags = "40"
mTradingClass = Value
End Property

Public Property Get TradingClass() As String
TradingClass = mTradingClass
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Equals(ByVal pContractSpecifier As IContractSpecifier) As Boolean
Const ProcName As String = "Equals"
On Error GoTo Err

Equals = GContractUtils.ContractSpecsEqual(Me, pContractSpecifier)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

'Public Function FuzzyEquals(ByVal pContractSpecifier As IContractSpecifier) As Boolean
'Const ProcName As String = "FuzzyEquals"
'
'On Error GoTo Err
'
'If pContractSpecifier Is Nothing Then Exit Function
'FuzzyEquals = (Me.SecType = pContractSpecifier.SecType And _
'                Me.Symbol = pContractSpecifier.Symbol) Or _
'            (Me.LocalSymbol = pContractSpecifier.LocalSymbol)
'
'Exit Function
'
'Err:
'GContracts.HandleUnexpectedError ProcName, ModuleName
'End Function

Public Sub Initialise( _
                Optional ByVal LocalSymbol As String, _
                Optional ByVal Symbol As String, _
                Optional ByVal TradingClass As String, _
                Optional ByVal Exchange As String, _
                Optional ByVal SecType As SecurityTypes = SecTypeNone, _
                Optional ByVal CurrencyCode As String, _
                Optional ByVal Expiry As String, _
                Optional ByVal Multiplier As Double = DefaultMultiplier, _
                Optional ByVal Strike As Double, _
                Optional ByVal Right As OptionRights = OptionRights.OptNone)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument LocalSymbol <> "" Or Symbol <> "", "Symbol must be supplied if LocalSymbol is not supplied"
AssertArgument Exchange = "" Or GContractUtils.IsValidExchangeSpecifier(Exchange), "'" & Exchange & "' is not a valid Exchange specifier"
Dim lErrorMessage As String
If Not GContractUtils.IsValidExpiry(Expiry, lErrorMessage) Then _
    AssertArgument False, "'" & Expiry & "' is not a valid Expiry format: " & lErrorMessage

Select Case SecType
Case 0  ' ie not supplied
Case SecTypeStock
Case SecTypeFuture
Case SecTypeCFD
Case SecTypeCrypto
Case SecTypeOption, SecTypeFuturesOption, SecTypeWarrant
    AssertArgument Strike >= 0, "Strike must be >= 0"
    Select Case Right
    Case OptCall
    Case OptPut
    Case OptNone
    Case Else
        AssertArgument False, "'" & Right & "' is not a valid option Right"
    End Select
Case SecTypeCash
Case SecTypeCombo
    AssertArgument False, "Sectype 'combo' is not permissible"
Case SecTypeIndex
Case Else
    AssertArgument False, "'" & SecType & "' is not a valid secType"
End Select

mLocalSymbol = LocalSymbol
mSymbol = Symbol
mTradingClass = TradingClass
mExchange = Exchange
mSecType = SecType
mCurrency = CurrencyCode
mExpiry = Expiry
mMultiplier = Multiplier
mStrike = Strike
mRight = Right

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ToJSON() As String
Const ProcName As String = "ToJSON"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeNameAndValue "ComboLegs", mComboLegs, Nothing

jb.EncodeName "zCurrencyCode"
jb.EncodeValue mCurrency

jb.EncodeName "zExchange"
jb.EncodeValue mExchange

jb.EncodeNameAndValue "zExpiry", mExpiry, DefaultExpiry

jb.EncodeName "zLocalSymbol"
jb.EncodeValue mLocalSymbol

jb.EncodeName "zMultiplier"
jb.EncodeValue mMultiplier

jb.EncodeNameAndValue "zRight", mRight, OptNone

jb.EncodeName "zSecType"
jb.EncodeValue mSecType

jb.EncodeNameAndValue "zStrike", mStrike, 0#

jb.EncodeName "zSymbol"
jb.EncodeValue mSymbol

jb.EncodeName "zTradingClass"
jb.EncodeValue mTradingClass

jb.EncodeNameAndValue "ProviderProperties", mProviderProperties, Nothing

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

Dim comboLegObj As ComboLeg
Dim i As Long

ToString = IIf(mLocalSymbol <> "", "[" & mLocalSymbol & "]", "")
ToString = ToString & IIf(mSecType <> 0, GContractUtils.SecTypeToShortString(mSecType) & "#", "")
ToString = ToString & mSymbol & IIf(mTradingClass <> "", "/" & mTradingClass, "")
Select Case mRight
Case OptCall, OptPut
    ToString = ToString & "=" & GContractUtils.OptionRightToString(mRight)
End Select
ToString = ToString & IIf(mStrike <> 0, Format(mStrike, "0.00"), "")
ToString = ToString & IIf(mExpiry <> "", "(" & mExpiry & ")", "")
ToString = ToString & IIf(mExchange <> "", "@" & mExchange, "")
ToString = ToString & IIf(mCurrency <> "", "$" & mCurrency, "")
ToString = ToString & IIf(mMultiplier <> 0, "'" & mMultiplier, "")

i = 0
If Not mComboLegs Is Nothing Then
    ToString = ToString & "{"
    For Each comboLegObj In mComboLegs
        If i > 0 Then ToString = ToString & ","
        i = i + 1
        ToString = ToString & comboLegObj.ToString
    Next
    ToString = ToString & "}"
End If

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

