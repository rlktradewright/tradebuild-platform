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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSymbol       As String
Private mSecType      As SecurityTypes
Private mExpiry       As String
Private mStrike       As Double
Private mRight        As OptionRights
Private mExchange     As String
Private mCurrency     As String
Private mLocalSymbol  As String
Private mComboLegs    As ComboLegs

Private mLocked As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

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

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ComboLegs(ByVal value As ComboLegs)
Set mComboLegs = value
End Property

Public Property Get ComboLegs() As ComboLegs
If mComboLegs Is Nothing Then
    Set mComboLegs = New ComboLegs
    If mLocked Then mComboLegs.locked = True
End If
Set ComboLegs = mComboLegs
End Property

Friend Property Let currencyCode(ByVal value As String)
mCurrency = UCase$(value)
End Property

Public Property Get currencyCode() As String
currencyCode = mCurrency
End Property

Friend Property Let exchange(ByVal value As String)
mExchange = UCase$(value)
End Property

Public Property Get exchange() As String
exchange = mExchange
End Property

Friend Property Let expiry(ByVal value As String)
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

Friend Property Let localSymbol(ByVal value As String)
mLocalSymbol = UCase$(value)
End Property

Public Property Get localSymbol() As String
localSymbol = mLocalSymbol
End Property

Friend Property Let locked(ByVal value As Boolean)
mLocked = value
If Not mComboLegs Is Nothing Then mComboLegs.locked = mLocked
End Property

Friend Property Get locked() As Boolean
locked = mLocked
End Property

Friend Property Let right(ByVal value As OptionRights)
mRight = value
End Property

Public Property Get right() As OptionRights
right = mRight
End Property

Friend Property Let sectype(ByVal value As SecurityTypes)
mSecType = value
End Property

Public Property Get sectype() As SecurityTypes
sectype = mSecType
End Property

Friend Property Let strike(ByVal value As Double)
mStrike = value
End Property

Public Property Get strike() As Double
strike = mStrike
End Property

Friend Property Let symbol(ByVal value As String)
mSymbol = UCase$(value)
End Property

Public Property Get symbol() As String
symbol = mSymbol
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function clone() As ContractSpecifier
Set clone = New ContractSpecifier
With clone
    If Not mComboLegs Is Nothing Then
        .ComboLegs = mComboLegs.clone
    End If
    .currencyCode = mCurrency
    .exchange = mExchange
    .expiry = mExpiry
    .localSymbol = mLocalSymbol
    .right = mRight
    .sectype = mSecType
    .strike = mStrike
    .symbol = mSymbol
    .locked = mLocked
End With
End Function

Public Function Equals(ByVal pContractSpecifier As ContractSpecifier) As Boolean
If pContractSpecifier Is Me Then Equals = True
Equals = (Me.key = pContractSpecifier.key)
End Function

Public Function FuyEquals(ByVal pContractSpecifier As ContractSpecifier) As Boolean
FuyEquals = (Me.sectype = pContractSpecifier.sectype And _
                Me.symbol = pContractSpecifier.symbol)
End Function

Public Function ToString() As String
Dim comboLegObj As ComboLeg
Dim i As Long

ToString = IIf(mLocalSymbol <> "", "localsymbol=" & mLocalSymbol & "; ", "")
ToString = ToString & IIf(mSymbol <> "", "symbol=" & mSymbol & "; ", "")
ToString = ToString & IIf(mSecType <> 0, "sectype=" & gSecTypeToString(mSecType) & "; ", "")
ToString = ToString & IIf(mExpiry <> "", "expiry=" & mExpiry & "; ", "")
ToString = ToString & IIf(mExchange <> "", "exchange=" & mExchange & "; ", "")
ToString = ToString & IIf(mCurrency <> "", "currencycode=" & mCurrency & "; ", "")
Select Case mRight
Case OptCall, OptPut
    ToString = ToString & "right=" & gOptionRightToString(mRight) & "; "
    ToString = ToString & "strike=" & mStrike & "; "
End Select

i = 0
If Not mComboLegs Is Nothing Then
    For Each comboLegObj In mComboLegs
        i = i + 1
        ToString = ToString & "Combo leg" & i & ":" & "; " & _
                    comboLegObj.ToString
    Next
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

