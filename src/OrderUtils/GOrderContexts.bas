Attribute VB_Name = "GOrderContexts"
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

Private Const ModuleName                            As String = "GOrderContexts"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderContextsCollection                    As New EnumerableCollection

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

Public Function gCreateOrderContexts( _
                ByVal pName As String, _
                ByVal pGroupName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pContractFuture As IFuture, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pOrderAuthoriser As IOrderAuthoriser, _
                ByVal pAccumulatedBracketOrders As BracketOrders, _
                ByVal pAccumulatedOrders As Orders, _
                ByVal pSimulatedClockFuture As IFuture) As OrderContexts
Const ProcName As String = "gCreateOrderContexts"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture is Nothing"
AssertArgument Not pOrderSubmitter Is Nothing, "pOrderSubmitter is Nothing"

Set gCreateOrderContexts = New OrderContexts
gCreateOrderContexts.Initialise pName, pGroupName, pIsSimulated, pContractFuture, pDataSource, pOrderSubmitter, pOrderAuthoriser, pAccumulatedBracketOrders, pAccumulatedOrders, pSimulatedClockFuture
mOrderContextsCollection.Add gCreateOrderContexts, pName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetOrderContexts( _
                ByVal pName As String) As OrderContexts
Const ProcName As String = "gGetOrderContexts"
On Error GoTo Err

If mOrderContextsCollection.Contains(pName) Then Set gGetOrderContexts = mOrderContextsCollection.Item(pName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gRemoveorderContexts(ByVal pOrderContexts As OrderContexts)
Const ProcName As String = "gRemoveorderContexts"
On Error GoTo Err

mOrderContextsCollection.Remove pOrderContexts.Name

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




