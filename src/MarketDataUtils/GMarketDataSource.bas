Attribute VB_Name = "GMarketDataSource"
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

Private Const ModuleName                            As String = "GMarketDataSource"

'@================================================================================
' Member variables
'@================================================================================

Private mHandleAllocator                            As New HandleAllocator

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

Public Function gGetCurrentTickSummary(ByVal pDataSource As IMarketDataSource) As String
Const ProcName As String = "gGetCurrentTickSummary"
On Error GoTo Err

Dim s As String

AssertArgument (pDataSource.State = MarketDataSourceStateReady Or _
                pDataSource.State = MarketDataSourceStateRunning), "Tick summary not available"

Dim lContract As IContract
Set lContract = pDataSource.ContractFuture.Value
Dim lSectype As SecurityTypes: lSectype = lContract.Specifier.SecType
Dim lTickSize As Double: lTickSize = lContract.TickSize

Dim lTick As GenericTick

s = "B=" & formatPriceTick(pDataSource, TickTypeBid, lSectype, lTickSize)
s = s & ";A=" & formatPriceTick(pDataSource, TickTypeAsk, lSectype, lTickSize)
s = s & ";T=" & formatPriceTick(pDataSource, TickTypeTrade, lSectype, lTickSize)
s = s & ";V=" & formatSizeTick(pDataSource, TickTypeVolume)

gGetCurrentTickSummary = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Methods
'@================================================================================

Public Function gAllocateHandle() As Long
Const ProcName As String = "gAllocateHandle"
On Error GoTo Err

gAllocateHandle = mHandleAllocator.AllocateHandle

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gReleaseHandle(ByVal pHandle As Long)
Const ProcName As String = "gReleaseHandle"
On Error GoTo Err

mHandleAllocator.ReleaseHandle pHandle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatPriceTick( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTickType As TickTypes, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double) As String
Const ProcName As String = "formatPriceTick"
On Error GoTo Err

Dim s As String

Dim lTick As GenericTick

If pDataSource.HasCurrentTick(pTickType) Then
    lTick = pDataSource.CurrentTick(pTickType)
    s = s & FormatPrice(lTick.Price, pSecType, pTickSize)
    s = s & "(" & lTick.Size
    s = s & ")"
Else
    s = s & "n/a"
End If

formatPriceTick = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function formatSizeTick( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTickType As TickTypes) As String
Const ProcName As String = "formatSizeTick"
On Error GoTo Err

If pDataSource.HasCurrentTick(pTickType) Then
    formatSizeTick = pDataSource.CurrentTick(pTickType).Size
Else
    formatSizeTick = "n/a"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




