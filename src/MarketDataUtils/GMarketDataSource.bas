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
s = gPadStringRight("B=" & formatPriceTick(pDataSource, TickTypeBid, lSectype, lTickSize), 17)
s = s & gPadStringRight("A=" & formatPriceTick(pDataSource, TickTypeAsk, lSectype, lTickSize), 17)
If lSectype <> SecTypeCash Then
    s = s & gPadStringRight("T=" & formatPriceTick(pDataSource, TickTypeTrade, lSectype, lTickSize), 17)
    If lSectype = SecTypeOption Or lSectype = SecTypeFuturesOption Then
        s = s & gPadStringRight("M=" & formatPriceTick(pDataSource, TickTypeOptionModelPrice, lSectype, lTickSize), 17)
    End If
    s = s & "V=" & formatSizeTick(pDataSource, TickTypeVolume)
End If

gGetCurrentTickSummary = s

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
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
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPadStringleft(ByRef pInput As String, ByVal pLength As Long) As String
Dim lInput As String: lInput = Right$(pInput, pLength)
gPadStringleft = Space$(pLength - Len(lInput)) & lInput
End Function

Public Function gPadStringRight(ByRef pInput As String, ByVal pLength As Long) As String
Dim lInput As String: lInput = Left$(pInput, pLength)
gPadStringRight = lInput & Space$(pLength - Len(lInput))
End Function

Public Sub gReleaseHandle(ByVal pHandle As Long)
Const ProcName As String = "gReleaseHandle"
On Error GoTo Err

mHandleAllocator.ReleaseHandle pHandle

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatBigInteger(ByVal n As Long) As String
Dim s As String
If n < 9950 Then
    s = CStr(n)
ElseIf n < 99500 Then
    s = Format((n / 1000), "0.0") & "K"
ElseIf n < 999500 Then
    s = Format((n / 1000), "0") & "K"
ElseIf n < 99950000 Then
    s = Format((n / 1000000), "0.0") & "M"
ElseIf n < 999500000 Then
    s = Format((n / 1000000), "0") & "M"
Else
    s = Format((n / 1000000000), "0.0") & "G"
End If


formatBigInteger = s
End Function

Private Function formatPriceTick( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTickType As TickTypes, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double) As String
Const ProcName As String = "formatPriceTick"
On Error GoTo Err

Dim s As String

If Not pDataSource.HasCurrentTick(pTickType) Then
    s = s & "n/a"
Else
    Dim lTick As GenericTick
    lTick = pDataSource.CurrentTick(pTickType)
    s = s & FormatPrice(lTick.Price, pSecType, pTickSize)
    If lTick.Size.EQ(DecimalZero) Then
        s = s & "n/a"
    Else
        s = s & "("
        s = s & formatBigInteger(lTick.Size)
        s = s & ")"
    End If
End If

formatPriceTick = s

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function formatSizeTick( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTickType As TickTypes) As String
Const ProcName As String = "formatSizeTick"
On Error GoTo Err

If pDataSource.HasCurrentTick(pTickType) Then
    formatSizeTick = formatBigInteger(pDataSource.CurrentTick(pTickType).Size)
Else
    formatSizeTick = "n/a"
End If

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function




