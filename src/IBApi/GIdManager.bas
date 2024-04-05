Attribute VB_Name = "GIdManager"
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

Public Enum IdTypes
    IdTypeNone
    IdTypeMarketData
    IdTypeMarketDepth
    IdTypeHistoricalData
    IdTypeOrder
    IdTypeContractData
    IdTypeExecution
    IdTypeAccount
    IdTypeScanner
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GIdManager"

Private Const BaseMarketDataRequestId               As Long = 0
Private Const BaseMarketDepthRequestId              As Long = &H40000
Private Const BaseScannerRequestId                  As Long = &H41000
Private Const BaseHistoricalDataRequestId           As Long = &H60000
Private Const BaseExecutionsRequestId               As Long = &HC0000
Private Const BaseContractRequestId                 As Long = &H100000
Private Const BaseAccountRequestId                  As Long = &H200000
Public Const BaseOrderId                            As Long = &H10000000

Public Const MaxCallersMarketDataRequestId          As Long = BaseMarketDepthRequestId - 1
Public Const MaxCallersMarketDepthRequestId         As Long = BaseScannerRequestId - BaseMarketDepthRequestId - 1
Public Const MaxCallersScannerRequestId             As Long = BaseHistoricalDataRequestId - BaseScannerRequestId - 1
Public Const MaxCallersHistoricalDataRequestId      As Long = BaseExecutionsRequestId - BaseHistoricalDataRequestId - 1
Public Const MaxCallersExecutionsRequestId          As Long = BaseContractRequestId - BaseExecutionsRequestId - 1
Public Const MaxCallersContractRequestId            As Long = BaseAccountRequestId - BaseContractRequestId - 1
Public Const MaxCallersAccountRequestId             As Long = BaseOrderId - BaseAccountRequestId - 1

'@================================================================================
' Member variables
'@================================================================================

Private mNextOrderID                                As Long

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

Public Function GetCallerId(ByVal pTwsId As Long, ByVal pIdType As IdTypes) As Long
Select Case pIdType
Case IdTypes.IdTypeMarketData
    GetCallerId = pTwsId - BaseMarketDataRequestId
Case IdTypes.IdTypeMarketDepth
    GetCallerId = pTwsId - BaseMarketDepthRequestId
Case IdTypes.IdTypeHistoricalData
    GetCallerId = pTwsId - BaseHistoricalDataRequestId
Case IdTypes.IdTypeExecution
    GetCallerId = pTwsId - BaseExecutionsRequestId
Case IdTypes.IdTypeContractData
    GetCallerId = pTwsId - BaseContractRequestId
Case IdTypes.IdTypeOrder
    GetCallerId = pTwsId - BaseOrderId
Case IdTypes.IdTypeScanner
    GetCallerId = pTwsId - BaseScannerRequestId
Case IdTypes.IdTypeAccount
    GetCallerId = pTwsId - BaseAccountRequestId
Case Else
    AssertArgument False, "invalid id type " & pIdType
End Select
End Function

Public Function GetIdType( _
                ByVal id As Long) As IdTypes
Const ProcName As String = "ggGetIdType"
On Error GoTo Err

If id >= BaseOrderId Then
    GetIdType = IdTypeOrder
ElseIf id >= BaseAccountRequestId Then
    GetIdType = IdTypeAccount
ElseIf id >= BaseContractRequestId Then
    GetIdType = IdTypeContractData
ElseIf id >= BaseExecutionsRequestId Then
    GetIdType = IdTypeExecution
ElseIf id >= BaseHistoricalDataRequestId Then
    GetIdType = IdTypeHistoricalData
ElseIf id >= BaseScannerRequestId Then
    GetIdType = IdTypeScanner
ElseIf id >= BaseMarketDepthRequestId Then
    GetIdType = IdTypeMarketDepth
ElseIf id >= 0 Then
    GetIdType = IdTypeMarketData
Else
    GetIdType = IdTypeNone
End If

Exit Function

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Public Function GetTwsId(ByVal pCallerId As Long, ByVal pIdType As IdTypes) As Long
AssertArgument pCallerId >= 0, "Id must not be negative"
Select Case pIdType
Case IdTypes.IdTypeMarketData
    AssertArgument pCallerId <= MaxCallersMarketDataRequestId, "Max request id is " & MaxCallersMarketDataRequestId
    GetTwsId = pCallerId + BaseMarketDataRequestId
Case IdTypes.IdTypeMarketDepth
    AssertArgument pCallerId <= MaxCallersMarketDepthRequestId, "Max request id is " & MaxCallersMarketDepthRequestId
    GetTwsId = pCallerId + BaseMarketDepthRequestId
Case IdTypes.IdTypeHistoricalData
    AssertArgument pCallerId <= MaxCallersHistoricalDataRequestId, "Max request id is " & MaxCallersHistoricalDataRequestId
    GetTwsId = pCallerId + BaseHistoricalDataRequestId
Case IdTypes.IdTypeScanner
    AssertArgument pCallerId <= MaxCallersScannerRequestId, "Max request id is " & MaxCallersScannerRequestId
    GetTwsId = pCallerId + BaseScannerRequestId
Case IdTypes.IdTypeExecution
    AssertArgument pCallerId <= MaxCallersExecutionsRequestId, "Max request id is " & MaxCallersExecutionsRequestId
    GetTwsId = pCallerId + BaseExecutionsRequestId
Case IdTypes.IdTypeContractData
    AssertArgument pCallerId <= MaxCallersContractRequestId, "Max request id is " & MaxCallersContractRequestId
    GetTwsId = pCallerId + BaseContractRequestId
Case IdTypes.IdTypeOrder
    AssertArgument False, "Invalid call"
Case Else
    AssertArgument False, "Invalid ID Type " & pIdType
End Select
End Function

Public Function GetNextOrderId(Optional ByVal pPeekOnly As Boolean = False) As Long
If mNextOrderID = 0 Then mNextOrderID = BaseOrderId
GetNextOrderId = mNextOrderID
If Not pPeekOnly Then mNextOrderID = mNextOrderID + 1
End Function

Public Sub SetNextOrderId(ByVal Value As Long)
AssertArgument Value > mNextOrderID, "Value must be >= " & mNextOrderID
mNextOrderID = Value
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




