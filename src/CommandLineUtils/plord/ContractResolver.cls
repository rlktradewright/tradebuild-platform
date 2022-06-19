VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "ContractResolver"

'@================================================================================
' Member variables
'@================================================================================

Private mContractFuture                             As IFuture
Attribute mContractFuture.VB_VarHelpID = -1

Private mContractStore                              As IContractStore

Private mContractSpec                               As IContractSpecifier

Private mSelectionMode                              As OptionStrikeSelectionModes
Private mParameter                                  As Long
Private mOperator                                   As OptionStrikeSelectionOperators
Private mUnderlyingExchangeName                     As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractName() As String
With mContractSpec
    Dim lStrikeStr As String
    If mSelectionMode = OptionStrikeSelectionModeNone Then
        If .Strike <> 0 Then lStrikeStr = CStr(.Strike)
    Else
        lStrikeStr = mParameter & gStrikeSelectionModeToString(mSelectionMode) & _
                    IIf(mUnderlyingExchangeName <> "", "," & mUnderlyingExchangeName, "")
    End If

    ContractName = IIf(.LocalSymbol <> "", "[" & .LocalSymbol & "]", "") & _
                    IIf(.SecType <> 0, SecTypeToShortString(.SecType) & ":", "") & _
                    .Symbol & _
                    IIf(.Right <> OptNone, "=" & Left$(OptionRightToString(.Right), 1), "") & _
                    lStrikeStr & _
                    IIf(.Expiry <> "", "(" & .Expiry & ")", "") & _
                    IIf(.Exchange <> "", "@" & .Exchange, "") & _
                    IIf(.CurrencyCode <> "", "$" & .CurrencyCode, "") & _
                    IIf(.Multiplier <> 0, "'" & .Multiplier, "")
End With
End Property

Friend Property Get StrikeSelectionMode() As OptionStrikeSelectionModes
StrikeSelectionMode = mSelectionMode
End Property

Friend Property Get StrikeSelectionOperator() As OptionStrikeSelectionOperators
StrikeSelectionOperator = mOperator
End Property

Friend Property Get StrikeSelectionParameter() As Long
StrikeSelectionParameter = mParameter
End Property

Friend Property Get UnderlyingExchangeName() As String
UnderlyingExchangeName = mUnderlyingExchangeName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mContractFuture.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ResolveContract( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractStore As IContractStore, _
                ByVal pSelectionMode As OptionStrikeSelectionModes, _
                ByVal pParameter As Long, _
                ByVal pOperator As OptionStrikeSelectionOperators, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractSpec = pContractSpec
Set mContractStore = pContractStore
mSelectionMode = pSelectionMode
mParameter = pParameter
mOperator = pOperator
mUnderlyingExchangeName = pUnderlyingExchangeName

Select Case mSelectionMode
Case OptionStrikeSelectionModeNone
    Set mContractFuture = FetchContract(pContractSpec, mContractStore)
Case OptionStrikeSelectionModeExpenditure, _
        OptionStrikeSelectionModeDelta
    gWriteLineToConsole "Resolving contract"
    Set mContractFuture = GetOptionContract( _
                            pContractSpec, _
                            OrderActionBuy, _
                            pContractStore, _
                            mSelectionMode, _
                            mParameter, _
                            mOperator, _
                            pUnderlyingExchangeName, _
                            pMarketDataManager)
Case Else
    Assert False, "Unexpected strike selection mode: " & pSelectionMode
End Select

Set ResolveContract = mContractFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================



