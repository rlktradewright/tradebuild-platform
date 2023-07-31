VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsContractSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TwsContractSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mConId                                      As Long
Private mSymbol                                     As String
Private mSecType                                    As TwsSecTypes
Private mExpiry                                     As String
Private mStrike                                     As Double
Private mOptRight                                   As TwsOptionRights
Private mMultiplier                                 As Double
Private mExchange                                   As String

Private mCurrencyCode                               As String
Private mLocalSymbol                                As String
Private mTradingClass                               As String
Private mPrimaryExch                                As String   ' pick a non-aggregate (ie not the SMART Exchange) Exchange that the contract trades on.  DO NOT SET TO SMART.

' COMBOS
Private mComboLegsDescription                       As String   ' received in open Order version 14 and up for all combos
Private mComboLegs                                  As New TwsComboLegs

' delta neutral
Private mDeltaNeutralContract                       As TwsDeltaNeutralContract

Private mAr()                                       As String
Private mArIndex                                    As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ComboLegs() As TwsComboLegs
Set ComboLegs = mComboLegs
End Property

Public Property Let ComboLegsDescription(ByVal Value As String)
mComboLegsDescription = Value
End Property

Public Property Get ComboLegsDescription() As String
ComboLegsDescription = mComboLegsDescription
End Property

Public Property Let ConId(ByVal Value As Long)
mConId = Value
End Property

Public Property Get ConId() As Long
ConId = mConId
End Property

Public Property Let CurrencyCode(ByVal Value As String)
mCurrencyCode = Value
End Property

Public Property Get CurrencyCode() As String
CurrencyCode = mCurrencyCode
End Property

Public Property Let Expiry(ByVal Value As String)
mExpiry = Value
End Property

Public Property Get Expiry() As String
Expiry = mExpiry
End Property

Public Property Let Exchange(ByVal Value As String)
mExchange = Value
End Property

Public Property Get Exchange() As String
Exchange = mExchange
End Property

Public Property Let LocalSymbol(ByVal Value As String)
mLocalSymbol = Value
End Property

Public Property Get LocalSymbol() As String
LocalSymbol = mLocalSymbol
End Property

Public Property Let Multiplier(ByVal Value As Double)
mMultiplier = Value
End Property

Public Property Get Multiplier() As Double
Multiplier = mMultiplier
End Property

Public Property Let OptRight(ByVal Value As TwsOptionRights)
mOptRight = Value
End Property

Public Property Get OptRight() As TwsOptionRights
OptRight = mOptRight
End Property

Public Property Let PrimaryExch(ByVal Value As String)
mPrimaryExch = Value
End Property

Public Property Get PrimaryExch() As String
PrimaryExch = mPrimaryExch
End Property

Public Property Let SecType(ByVal Value As TwsSecTypes)
mSecType = Value
End Property

Public Property Get SecType() As TwsSecTypes
SecType = mSecType
End Property

Public Property Let Strike(ByVal Value As Double)
mStrike = Value
End Property

Public Property Get Strike() As Double
Strike = mStrike
End Property

Public Property Let Symbol(ByVal Value As String)
mSymbol = Value
End Property

Public Property Get Symbol() As String
Symbol = mSymbol
End Property

Public Property Let TradingClass(ByVal Value As String)
mTradingClass = Value
End Property

Public Property Get TradingClass() As String
TradingClass = mTradingClass
End Property

Public Property Let DeltaNeutralContract(ByRef Value As TwsDeltaNeutralContract)
mDeltaNeutralContract = Value
End Property

Public Property Get DeltaNeutralContract() As TwsDeltaNeutralContract
DeltaNeutralContract = mDeltaNeutralContract
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

ReDim mAr(3) As String
mArIndex = 0

addField "Local symbol", mLocalSymbol
addField "Exchange", mExchange
addField "Sec type", gTwsSecTypeToShortString(mSecType)
addField "Expiry", mExpiry
addField "Trading class", mTradingClass
addField "Con id", CStr(mConId)
addField "Symbol", mSymbol
addField "Currency", mCurrencyCode
If mStrike <> 0# Then addField "Strike", mStrike
addField "Right", gTwsOptionRightToString(mOptRight)
addField "Multiplier", mMultiplier
addField "Primary exch", mPrimaryExch

If mComboLegs.Count > 0 Then
    addField "Como legs count", mComboLegs.Count
    Dim lLeg As TwsComboLeg
    Dim i As Long
    For Each lLeg In mComboLegs
        i = i + 1
        addField "Con id " & i, lLeg.ConId
        addField "Exchange " & i, lLeg.Exchange
        addField "Action " & i, gTwsOrderActionToString(lLeg.Action)
        addField "Ratio " & i, lLeg.Ratio
        addField "Open/close " & i, IIf(lLeg.OpenClose = TwsLegOpenCloseSame, "", CStr(lLeg.OpenClose))
        addField "Designated location " & i, lLeg.DesignatedLocation
        addField "Short sale slot " & i, IIf(lLeg.ShortSaleSlot = TwsShortSaleSlotNotApplicable, "", CStr(lLeg.ShortSaleSlot))
    Next
End If

If mDeltaNeutralContract.ConId <> 0 Then
    addField "Delta neutral con id", mDeltaNeutralContract.ConId
    addField "Delta neutral delta", mDeltaNeutralContract.Delta
    addField "Delta neutral price", mDeltaNeutralContract.Price
End If

ReDim Preserve mAr(mArIndex - 1) As String
ToString = Join(mAr, "")

Erase mAr

Exit Function

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addField(ByVal pName As String, ByVal pValue As String)
If pValue = "" Then Exit Sub
If mArIndex <> 0 Then addItem "; "
addItem pName
addItem ": "
addItem pValue
End Sub

Private Sub addItem(ByVal pItem As String)
If mArIndex > UBound(mAr) Then ReDim Preserve mAr(2 * (UBound(mAr) + 1) - 1) As String
mAr(mArIndex) = pItem
mArIndex = mArIndex + 1
End Sub



