Attribute VB_Name = "GTimeframeUtils"
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

Private Const ModuleName                            As String = "GTimeframeUtils"

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

Public Function CreateTimeframes( _
                ByVal pStudyBase As IStudyBase, _
                Optional ByVal pContractFuture As IFuture, _
                Optional ByVal pHistDataStore As IHistoricalDataStore, _
                Optional ByVal pClockFuture As IFuture, _
                Optional ByVal pBarType As BarTypes = BarTypeTrade) As Timeframes
Const ProcName As String = "CreateTimeframes"
On Error GoTo Err

If Not pHistDataStore Is Nothing Then
    Select Case pBarType
    Case BarTypeTrade
        AssertArgument pHistDataStore.Supports(HistDataStoreCapabilityFetchTradeBars), "Cannot fetch historical trade bars"
    Case BarTypeBid, BarTypeAsk
        AssertArgument pHistDataStore.Supports(HistDataStoreCapabilityFetchBidAndAskBars), "Cannot fetch historical bid and ask bars"
    Case Else
        AssertArgument False, "Invalid bar type"
    End Select
End If

Set CreateTimeframes = New Timeframes
CreateTimeframes.Initialise pStudyBase, pContractFuture, pHistDataStore, pClockFuture, pBarType

Exit Function

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================






