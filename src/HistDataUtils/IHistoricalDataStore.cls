VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IHistoricalDataStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Public Enum HistDataStoreCapabilities
    HistDataStoreCapabilityNone = 0
    HistDataStoreCapabilityWriteTradeBars = 1
    HistDataStoreCapabilityWriteBidAndAskBars = 2
    HistDataStoreCapabilityFetchTradeBars = 4
    HistDataStoreCapabilityFetchBidAndAskBars = 8
    HistDataStoreCapabilityWriteAsync = 16
    HistDataStoreCapabilityMarketScan = 32
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "IHistoricalDataStore"

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

Public Property Get Capabilities() As HistDataStoreCapabilities
End Property

Public Property Get TimePeriodValidator() As ITimePeriodValidator
End Property

Public Property Get WriterTimeframes() As TimePeriod()
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub CancelScanRefresh( _
                ByVal pRefreshCookie As String)
End Sub

Public Function CreateBarWriter( _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pContractFuture As IFuture) As IBarWriter
End Function

Public Function FetchBars( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarDataSpecifier As BarDataSpecifier, _
                Optional ByVal pSpecialFromTime As String, _
                Optional ByVal pSpecialToTime As String, _
                Optional ByVal pClockFuture As IFuture, _
                Optional ByVal pListener As IBarFetchListener, _
                Optional ByVal pDontReturnBars As Boolean, _
                Optional ByVal pCookie As Variant) As IFuture
End Function

Public Sub Finish()
End Sub

Public Function RefreshScanData( _
                ByVal pRefreshCookie As String) As IFuture
End Function

Public Function ScanMarket( _
                ByVal pScanName As String, _
                ByVal pScanParameters As Parameters, _
                Optional ByVal pOptions As Parameters, _
                Optional ByVal pFilters As Parameters, _
                Optional ByVal pAutoRefresh As Boolean, _
                Optional ByVal pRefreshCookie As String) As IFuture
End Function

Public Function Supports(ByVal pCapability As HistDataStoreCapabilities) As Boolean
End Function

'@================================================================================
' Helper Functions
'@================================================================================






