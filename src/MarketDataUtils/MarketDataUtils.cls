VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketDataUtils"
Attribute VB_GlobalNameSpace = True
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

Public Enum MarketDataErrorCodes
    MarketDataErrCancelled = vbObjectError + 1536
    MarketDataErrContractNotUnique
    MarketDataErrNoContract
    MarketDataErrContractFetchError
    MarketDataErrContractExpired
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MarketDataUtils"

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

Public Function CreateRealtimeDataManager( _
                ByVal pMarketDataFactory As IMarketDataFactory, _
                Optional ByVal pOptions As MarketDataSourceOptions = MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pDefaultStateChangeListener As StateChangeListener, _
                Optional ByVal pNumberOfMarketDepthRows As Long = 20) As IMarketDataManager
Const ProcName As String = "CreateRealtimeDataManager"
On Error GoTo Err

AssertArgument Not pMarketDataFactory Is Nothing, "pMarketDataFactory is Nothing"

Dim rtm As New RealTimeDataManager
rtm.Initialise pMarketDataFactory, pOptions, pDefaultStateChangeListener, pNumberOfMarketDepthRows
Set CreateRealtimeDataManager = rtm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickDataManager( _
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pOptions As MarketDataSourceOptions = MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pDefaultStateChangeListener As StateChangeListener, _
                Optional ByVal pNumberOfMarketDepthRows As Long = 20, _
                Optional ByVal pReplaySpeed As Long = 1, _
                Optional ByVal pReplayProgressEventInterval As Long = 1000, _
                Optional ByVal pTimestampAdjustmentStart As Double, _
                Optional ByVal pTimestampAdjustmentEnd As Double) As IMarketDataManager
Const ProcName As String = "CreateTickDataManager"
On Error GoTo Err

Dim lTickDataManager As New TickfileDataManager
lTickDataManager.Initialise pTickfileSpecifiers, pTickfileStore, pPrimaryContractStore, pSecondaryContractStore, pOptions, pDefaultStateChangeListener, pNumberOfMarketDepthRows, pReplaySpeed, pReplayProgressEventInterval, pTimestampAdjustmentStart, pTimestampAdjustmentEnd

Set CreateTickDataManager = lTickDataManager

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetFormattedPriceFromQuoteEvent(ByRef ev As QuoteEventData) As String
Dim lDataSource As IMarketDataSource
Const ProcName As String = "GetFormattedPriceFromQuoteEvent"
On Error GoTo Err

Set lDataSource = ev.Source
Dim lContract As IContract
Set lContract = lDataSource.ContractFuture.Value
GetFormattedPriceFromQuoteEvent = FormatPrice(ev.Quote.Price, lContract.Specifier.SecType, lContract.TickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




