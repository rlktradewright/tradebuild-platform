Attribute VB_Name = "GPositionManager"
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

Private Const ModuleName                            As String = "GPositionManager"

'@================================================================================
' Member variables
'@================================================================================

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

Public Function gCreatePositionManager( _
                ByVal pName As String, _
                ByVal pContractFuture As IFuture, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pScopeName As String, _
                ByVal pGroupName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pMoneyManager As IMoneyManager) As PositionManager
Const ProcName As String = "gCreatePositionManager"
On Error GoTo Err

Dim lClr As BracketOrderRecoveryCtlr
If pScopeName <> "" Then Set lClr = gGetBracketOrderRecoveryController(pScopeName)

Dim lPm As New PositionManager
lPm.Initialise pName, pContractFuture, pOrderSubmitter, pDataSource, lClr, pGroupName, pIsSimulated, pMoneyManager

Set gCreatePositionManager = lPm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gNextApplicationIndex() As Long
Static lNextApplicationIndex As Long

Const ProcName As String = "gNextApplicationIndex"

On Error GoTo Err

gNextApplicationIndex = lNextApplicationIndex
lNextApplicationIndex = lNextApplicationIndex + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





