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

Private mSymbol       As String
Private mSecType      As SecurityTypes
Private mExpiry       As String
Private mStrike       As Double
Private mRight        As OptionRights
Private mExchange     As String
Private mCurrency     As String
Private mLocalSymbol  As String
Private mComboLegs    As ComboLegs

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ComboLegs(ByVal value As ComboLegs)
Set mComboLegs = value
End Property

Public Property Get ComboLegs() As ComboLegs
Set ComboLegs = mComboLegs
End Property

Friend Property Let CurrencyCode(ByVal value As String)
mCurrency = UCase$(value)
End Property

Public Property Get CurrencyCode() As String
CurrencyCode = mCurrency
End Property

Friend Property Let Exchange(ByVal value As String)
mExchange = UCase$(value)
End Property

Public Property Get Exchange() As String
Exchange = mExchange
End Property

Friend Property Let Expiry(ByVal value As String)
Const ProcName As String = "Expiry"
Dim failpoint As String
On Error GoTo Err

If value = "" Then Exit Property
If Not gIsValidExpiry(value) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid Expiry"
End If

mExpiry = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Expiry() As String
Expiry = mExpiry
End Property

Public Property Get Key() As String
Const ProcName As String = "Key"
Dim failpoint As String
On Error GoTo Err

Key = mLocalSymbol & "|" & _
    CStr(mSecType) & "|" & _
    mSymbol & "|" & _
    Left$(mExpiry, 6) & "|" & _
    mStrike & "|" & _
    CStr(mRight) & "|" & _
    mExchange & "|" & _
    mCurrency

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let LocalSymbol(ByVal value As String)
mLocalSymbol = UCase$(value)
End Property

Public Property Get LocalSymbol() As String
LocalSymbol = mLocalSymbol
End Property

Friend Property Let Right(ByVal value As OptionRights)
mRight = value
End Property

Public Property Get Right() As OptionRights
Right = mRight
End Property

Friend Property Let SecType(ByVal value As SecurityTypes)
mSecType = value
End Property

Public Property Get SecType() As SecurityTypes
SecType = mSecType
End Property

Friend Property Let Strike(ByVal value As Double)
mStrike = value
End Property

Public Property Get Strike() As Double
Strike = mStrike
End Property

Friend Property Let Symbol(ByVal value As String)
mSymbol = UCase$(value)
End Property

Public Property Get Symbol() As String
Symbol = mSymbol
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Equals(ByVal pContractSpecifier As ContractSpecifier) As Boolean
Const ProcName As String = "Equals"
Dim failpoint As String
On Error GoTo Err

If pContractSpecifier Is Nothing Then Exit Function
If pContractSpecifier Is Me Then Equals = True
Equals = (Me.Key = pContractSpecifier.Key)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FuzzyEquals(ByVal pContractSpecifier As ContractSpecifier) As Boolean
Const ProcName As String = "FuzzyEquals"
Dim failpoint As String
On Error GoTo Err

If pContractSpecifier Is Nothing Then Exit Function
FuzzyEquals = (Me.SecType = pContractSpecifier.SecType And _
                Me.Symbol = pContractSpecifier.Symbol)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function ToString() As String
Dim comboLegObj As ComboLeg
Dim i As Long

Const ProcName As String = "ToString"
Dim failpoint As String
On Error GoTo Err

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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================

