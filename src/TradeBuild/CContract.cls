VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
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

Implements TradeBuildSP.IContractSpecifier

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

Private mSymbol       As String
Private mSecType      As SecurityTypes
Private mExpiry       As String
Private mStrike       As Double
Private mRight        As OptionRights
Private mExchange     As String
Private mCurrency     As String
Private mLocalSymbol  As String
Private mComboLegs    As ComboLegs

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_ReadProperties(PropBag As PropertyBag)
mSymbol = PropBag.ReadProperty("symbol")
mSecType = PropBag.ReadProperty("secType")
mExpiry = PropBag.ReadProperty("expiry")
mStrike = PropBag.ReadProperty("strike")
mRight = PropBag.ReadProperty("right")
mExchange = PropBag.ReadProperty("exchange")
mCurrency = PropBag.ReadProperty("curency")
mLocalSymbol = PropBag.ReadProperty("localSymbol")
On Error Resume Next
Set mComboLegs = PropBag.ReadProperty("combolegs")
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "symbol", mSymbol
PropBag.WriteProperty "secType", mSecType
PropBag.WriteProperty "expiry", mExpiry
PropBag.WriteProperty "strike", mStrike
PropBag.WriteProperty "right", mRight
PropBag.WriteProperty "exchange", mExchange
PropBag.WriteProperty "curency", mCurrency
PropBag.WriteProperty "localSymbol", mLocalSymbol
PropBag.WriteProperty "combolegs", mComboLegs
End Sub

'================================================================================
' IContractSpecifier Interface Members
'================================================================================

Private Function IContractSpecifier_Clone() As TradeBuildSP.IContractSpecifier
Set IContractSpecifier_Clone = clone
End Function

Private Property Let IContractSpecifier_ComboLegs(ByVal RHS As TradeBuildSP.IComboLegs)
ComboLegs = RHS
End Property

Private Property Get IContractSpecifier_ComboLegs() As TradeBuildSP.IComboLegs
Set IContractSpecifier_ComboLegs = ComboLegs
End Property

Private Property Let IContractSpecifier_CurrencyCode(ByVal RHS As String)
currencyCode = RHS
End Property

Private Property Get IContractSpecifier_CurrencyCode() As String
IContractSpecifier_CurrencyCode = currencyCode
End Property

Private Function IContractSpecifier_Equals( _
                            ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier) As Boolean
IContractSpecifier_Equals = Equals(pContractSpecifier)
End Function

Private Property Let IContractSpecifier_Exchange(ByVal RHS As String)
exchange = RHS
End Property

Private Property Get IContractSpecifier_Exchange() As String
IContractSpecifier_Exchange = exchange
End Property

Private Property Let IContractSpecifier_Expiry(ByVal RHS As String)
expiry = RHS
End Property

Private Property Get IContractSpecifier_Expiry() As String
IContractSpecifier_Expiry = expiry
End Property

Private Function IContractSpecifier_FuzzyEquals( _
                            ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier) As Boolean
IContractSpecifier_FuzzyEquals = Equals(pContractSpecifier)
End Function

Private Property Get IContractSpecifier_Key() As String
IContractSpecifier_Key = key
End Property

Private Property Let IContractSpecifier_LocalSymbol(ByVal RHS As String)
localSymbol = RHS
End Property

Private Property Get IContractSpecifier_LocalSymbol() As String
IContractSpecifier_LocalSymbol = localSymbol
End Property

Private Property Let IContractSpecifier_Right(ByVal RHS As TradeBuildSP.OptionRights)
right = RHS
End Property

Private Property Get IContractSpecifier_Right() As TradeBuildSP.OptionRights
IContractSpecifier_Right = right
End Property

Private Property Let IContractSpecifier_SecType(ByVal RHS As TradeBuildSP.SecurityTypes)
secType = RHS
End Property

Private Property Get IContractSpecifier_SecType() As TradeBuildSP.SecurityTypes
IContractSpecifier_SecType = secType
End Property

Private Property Let IContractSpecifier_Strike(ByVal RHS As Double)
strike = RHS
End Property

Private Property Get IContractSpecifier_Strike() As Double
IContractSpecifier_Strike = strike
End Property

Private Property Let IContractSpecifier_Symbol(ByVal RHS As String)
symbol = RHS
End Property

Private Property Get IContractSpecifier_Symbol() As String
IContractSpecifier_Symbol = symbol
End Property

Private Function IContractSpecifier_ToString() As String
IContractSpecifier_ToString = ToString
End Function

'================================================================================
' Properties
'================================================================================

Public Property Get ComboLegs() As ComboLegs
If mComboLegs Is Nothing Then
    Set mComboLegs = New ComboLegs
End If
Set ComboLegs = mComboLegs
End Property

Public Property Let currencyCode(ByVal value As String)
mCurrency = UCase$(value)
End Property

Public Property Get currencyCode() As String
currencyCode = mCurrency
End Property

Public Property Let exchange(ByVal value As String)
mExchange = UCase$(value)
End Property

Public Property Get exchange() As String
exchange = mExchange
End Property

Public Property Let expiry(ByVal value As String)
mExpiry = value
End Property

Public Property Get expiry() As String
expiry = mExpiry
End Property

Public Property Get key() As String
key = mLocalSymbol & "|" & _
    CStr(mSecType) & "|" & _
    mSymbol & "|" & _
    Left$(mExpiry, 6) & "|" & _
    mStrike & "|" & _
    CStr(mRight) & "|" & _
    mExchange & "|" & _
    mCurrency
End Property

Public Property Let localSymbol(ByVal value As String)
mLocalSymbol = UCase$(value)
End Property

Public Property Get localSymbol() As String
localSymbol = mLocalSymbol
End Property

Public Property Let right(ByVal value As OptionRights)
mRight = value
End Property

Public Property Get right() As OptionRights
right = mRight
End Property

Public Property Let secType(ByVal value As SecurityTypes)
mSecType = value
End Property

Public Property Get secType() As SecurityTypes
secType = mSecType
End Property

Public Property Let strike(ByVal value As Double)
mStrike = value
End Property

Public Property Get strike() As Double
strike = mStrike
End Property

Public Property Let symbol(ByVal value As String)
mSymbol = UCase$(value)
End Property

Public Property Get symbol() As String
symbol = mSymbol
End Property

'================================================================================
' Methods
'================================================================================

Public Function clone() As contractSpecifier
Set clone = New contractSpecifier
With clone
    If Not mComboLegs Is Nothing Then
        .ComboLegs = mComboLegs.clone
    End If
    .currencyCode = mCurrency
    .exchange = mExchange
    .expiry = mExpiry
    .localSymbol = mLocalSymbol
    .right = mRight
    .secType = mSecType
    .strike = mStrike
    .symbol = mSymbol
End With
End Function

Public Function Equals(ByVal pContractSpecifier As contractSpecifier) As Boolean
Equals = (Me.key = pContractSpecifier.key)
End Function

Public Function FuzzyEquals(ByVal pContractSpecifier As contractSpecifier) As Boolean
FuzzyEquals = (Me.secType = pContractSpecifier.secType And _
                Me.symbol = pContractSpecifier.symbol)
End Function

Public Function ToString() As String
Dim comboLegObj As ComboLeg
Dim i As Long

ToString = IIf(mLocalSymbol <> "", "localsymbol=" & mLocalSymbol & vbCrLf, "")
ToString = ToString & "symbol=" & mSymbol & vbCrLf & _
        "sectype="
Select Case mSecType
Case SecTypeStock
    ToString = ToString & "SecTypeStock" & vbCrLf
Case SecTypeFuture
    ToString = ToString & "SecTypeFuture" & vbCrLf
Case SecTypeOption
    ToString = ToString & "SecTypeOption" & vbCrLf
Case SecTypeFuturesOption
    ToString = ToString & "SecTypeFuturesOption" & vbCrLf
Case SecTypeCash
    ToString = ToString & "SecTypeCash" & vbCrLf
Case SecTypeBag
    ToString = ToString & "SecTypeBag" & vbCrLf
End Select
ToString = ToString & _
        IIf(mExpiry <> "", "expiry=" & mExpiry & vbCrLf, "") & _
        IIf(mExchange <> "", "exchange=" & mExchange & vbCrLf, "") & _
        IIf(mCurrency <> "", "currencycode=" & mCurrency & vbCrLf, "")
                
Select Case mRight
Case OptNone
Case OptCall
    ToString = ToString & "right=Call" & vbCrLf
    ToString = ToString & "strike=" & mStrike & vbCrLf
Case OptPut
    ToString = ToString & "right=Put" & vbCrLf
    ToString = ToString & "strike=" & mStrike & vbCrLf
End Select

i = 0
If Not mComboLegs Is Nothing Then
    For Each comboLegObj In mComboLegs
        i = i + 1
        ToString = ToString & "Combo leg" & i & ":" & vbCrLf & _
                    comboLegObj.ToString
    Next
End If
End Function

'================================================================================
' Helper Functions
'================================================================================

