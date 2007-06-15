Attribute VB_Name = "Globals"
Option Explicit

''
' Description here
'
' @remarks
' @see
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

Private Const ProjectName                   As String = "ContractUtils26"
Private Const ModuleName                    As String = "Globals"

'@================================================================================
' Member variables
'@================================================================================

Private mExchangeCodes() As String
Private mMaxExchangeCodesIndex As Long

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

'@================================================================================
' Methods
'@================================================================================

Public Function gGetExchangeCodes() As String()
If mMaxExchangeCodesIndex = 0 Then setupExchangeCodes
gGetExchangeCodes = mExchangeCodes
End Function

Public Function gIsValidExchangeCode(ByVal code As String) As Boolean
Dim bottom As Long
Dim top As Long
Dim middle As Long

If mMaxExchangeCodesIndex = 0 Then setupExchangeCodes

code = UCase$(code)
bottom = 0
top = mMaxExchangeCodesIndex
middle = Fix((bottom + top) / 2)

Do
    If code < mExchangeCodes(middle) Then
        top = middle
    ElseIf code > mExchangeCodes(middle) Then
        bottom = middle
    Else
        gIsValidExchangeCode = True
        Exit Function
    End If
    middle = Fix((bottom + top) / 2)
Loop Until bottom = middle

If code = mExchangeCodes(middle) Then gIsValidExchangeCode = True
End Function

Public Function gOptionRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case ""
    gOptionRightFromString = OptNone
Case "CALL"
    gOptionRightFromString = OptCall
Case "PUT"
    gOptionRightFromString = OptPut
End Select
End Function

Public Function gOptionRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptNone
    gOptionRightToString = ""
Case OptCall
    gOptionRightToString = "Call"
Case OptPut
    gOptionRightToString = "Put"
End Select
End Function

Public Function gSecTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "STOCK", "STK"
    gSecTypeFromString = SecTypeStock
Case "FUTURE", "FUT"
    gSecTypeFromString = SecTypeFuture
Case "OPTION", "OPT"
    gSecTypeFromString = SecTypeOption
Case "FUTURES OPTION", "FOP"
    gSecTypeFromString = SecTypeFuturesOption
Case "CASH"
    gSecTypeFromString = SecTypeCash
Case "COMBO", "CMB"
    gSecTypeFromString = SecTypeCombo
Case "INDEX", "IND"
    gSecTypeFromString = SecTypeIndex
End Select
End Function

Public Function gSecTypeToString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    gSecTypeToString = "Stock"
Case SecTypeFuture
    gSecTypeToString = "Future"
Case SecTypeOption
    gSecTypeToString = "Option"
Case SecTypeFuturesOption
    gSecTypeToString = "Futures Option"
Case SecTypeCash
    gSecTypeToString = "Cash"
Case SecTypeCombo
    gSecTypeToString = "Combo"
Case SecTypeIndex
    gSecTypeToString = "Index"
End Select
End Function

Public Function gSecTypeToShortString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    gSecTypeToShortString = "STK"
Case SecTypeFuture
    gSecTypeToShortString = "FUT"
Case SecTypeOption
    gSecTypeToShortString = "OPT"
Case SecTypeFuturesOption
    gSecTypeToShortString = "FOP"
Case SecTypeCash
    gSecTypeToShortString = "CASH"
Case SecTypeCombo
    gSecTypeToShortString = "CMB"
Case SecTypeIndex
    gSecTypeToShortString = "IND"
End Select
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addExchangeCode(ByVal code As String)
mMaxExchangeCodesIndex = mMaxExchangeCodesIndex + 1
If mMaxExchangeCodesIndex > UBound(mExchangeCodes) Then
    ReDim Preserve mExchangeCodes(UBound(mExchangeCodes) + 10) As String
End If
mExchangeCodes(mMaxExchangeCodesIndex) = UCase$(code)
End Sub

Private Sub setupExchangeCodes()
ReDim mExchangeCodes(100) As String
mMaxExchangeCodesIndex = -1

addExchangeCode "ACE"
addExchangeCode "AEB"
addExchangeCode "AMEX"
addExchangeCode "ARCA"

addExchangeCode "BELFOX"
addExchangeCode "BOX"
addExchangeCode "BRUT"
addExchangeCode "BTRADE"
addExchangeCode "BVME"

addExchangeCode "CAES"
addExchangeCode "CBOE"
addExchangeCode "CDE"
addExchangeCode "CFE"

addExchangeCode "DTB"

addExchangeCode "EBS"
addExchangeCode "ECBOT"
addExchangeCode "EUREX"
addExchangeCode "EUREXUS"

addExchangeCode "FTA"
addExchangeCode "FWB"

addExchangeCode "GLOBEX"

addExchangeCode "HKFE"

addExchangeCode "IBIS"
addExchangeCode "IDEAL"
addExchangeCode "IDEALPRO"
addExchangeCode "IDEM"
addExchangeCode "INET"
addExchangeCode "INSTINET"
addExchangeCode "ISE"
addExchangeCode "ISLAND"

addExchangeCode "LIFFE"
addExchangeCode "LIFFE_NF"
addExchangeCode "LSE"

addExchangeCode "MATIF"
addExchangeCode "MEFF"
addExchangeCode "MEFFRV"
addExchangeCode "MONEP"
addExchangeCode "MXT"

addExchangeCode "NASDAQ"
addExchangeCode "NQLX"
addExchangeCode "NYMEX"
addExchangeCode "NYSE"

addExchangeCode "OMS"
addExchangeCode "ONE"
addExchangeCode "OSE.JPN"

addExchangeCode "PHLX"
addExchangeCode "PINK"
addExchangeCode "PSE"

addExchangeCode "RDBK"

addExchangeCode "SBF"
addExchangeCode "SFB"
addExchangeCode "SGX"
addExchangeCode "SMART"
addExchangeCode "SNFE"
addExchangeCode "SOFFEX"
addExchangeCode "SWB"
addExchangeCode "SWX"

addExchangeCode "TSE"
addExchangeCode "TSE.JPN"

addExchangeCode "VENTURE"
addExchangeCode "VIRTX"
addExchangeCode "VWAP"

ReDim Preserve mExchangeCodes(mMaxExchangeCodesIndex) As String

End Sub
