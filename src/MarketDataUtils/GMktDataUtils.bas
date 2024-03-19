Attribute VB_Name = "GMktDataUtils"
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

Private Const ModuleName                            As String = "GMktDataUtils"

Public Const NullIndex                              As Long = -1

Public Const ConfigSectionContract                  As String = "Contract"

Public Const OneSecond                              As Double = 1# / 86400#
Public Const OneMillisec                            As Double = 1# / 86400# / 1000#

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
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pStudyLibManager As StudyLibraryManager, _
                Optional ByVal pOptions As MarketDataSourceOptions = MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pDefaultStateChangeListener As IStateChangeListener, _
                Optional ByVal pNumberOfMarketDepthRows As Long = 20) As IMarketDataManager
Const ProcName As String = "CreateRealtimeDataManager"
On Error GoTo Err

AssertArgument Not pMarketDataFactory Is Nothing, "pMarketDataFactory is Nothing"

Dim rtm As New RealTimeDataManager
rtm.Initialise pMarketDataFactory, _
                pPrimaryContractStore, _
                pSecondaryContractStore, _
                pStudyLibManager, _
                pOptions, _
                pDefaultStateChangeListener, _
                pNumberOfMarketDepthRows
Set CreateRealtimeDataManager = rtm

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSequentialTickDataManager( _
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pStudyLibManager As StudyLibraryManager, _
                Optional ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pOptions As MarketDataSourceOptions = MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pDefaultStateChangeListener As IStateChangeListener, _
                Optional ByVal pNumberOfMarketDepthRows As Long = 20, _
                Optional ByVal pReplaySpeed As Long = 1, _
                Optional ByVal pReplayProgressEventInterval As Long = 1000, _
                Optional ByVal pTimestampAdjustmentStart As Double, _
                Optional ByVal pTimestampAdjustmentEnd As Double) As IMarketDataManager
Const ProcName As String = "CreateSequentialTickDataManager"
On Error GoTo Err

Dim lTickDataManager As New TickfileDataManager
lTickDataManager.Initialise pTickfileSpecifiers, pTickfileStore, True, pStudyLibManager, pPrimaryContractStore, pSecondaryContractStore, pOptions, pDefaultStateChangeListener, pNumberOfMarketDepthRows, pReplaySpeed, pReplayProgressEventInterval, pTimestampAdjustmentStart, pTimestampAdjustmentEnd

Set CreateSequentialTickDataManager = lTickDataManager

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickDataManager( _
                ByVal pTickfileSpecifiers As TickFileSpecifiers, _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pStudyLibManager As StudyLibraryManager, _
                Optional ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pOptions As MarketDataSourceOptions = MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pDefaultStateChangeListener As IStateChangeListener, _
                Optional ByVal pNumberOfMarketDepthRows As Long = 20, _
                Optional ByVal pReplaySpeed As Long = 1, _
                Optional ByVal pReplayProgressEventInterval As Long = 1000, _
                Optional ByVal pTimestampAdjustmentStart As Double, _
                Optional ByVal pTimestampAdjustmentEnd As Double) As IMarketDataManager
Const ProcName As String = "CreateTickDataManager"
On Error GoTo Err

Dim lTickDataManager As New TickfileDataManager
lTickDataManager.Initialise pTickfileSpecifiers, pTickfileStore, False, pStudyLibManager, pPrimaryContractStore, pSecondaryContractStore, pOptions, pDefaultStateChangeListener, pNumberOfMarketDepthRows, pReplaySpeed, pReplayProgressEventInterval, pTimestampAdjustmentStart, pTimestampAdjustmentEnd

Set CreateTickDataManager = lTickDataManager

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrentTickSummary(ByVal pDataSource As IMarketDataSource) As String
Const ProcName As String = "GetCurrentTickSummary"
On Error GoTo Err

GetCurrentTickSummary = gGetCurrentTickSummary(pDataSource)

Exit Function

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
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
GMktData.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================







