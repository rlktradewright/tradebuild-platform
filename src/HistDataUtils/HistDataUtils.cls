VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataUtils"
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

Public Enum BarTypes
    BarTypeTrade = 0
    BarTypeBid = 1
    BarTypeAsk = 2
End Enum

Public Enum HistDataEventCodes
    HistDataEventFileDoesNotExist = 1
    HistDataEventFileIsEmpty
    HistDataEventFileIsInvalid
    HistDataEventFileFormatNotSupported
    HistDataEventNoContractDetails
    HistDataEventDataSourceNotAvailable
    HistDataEventAmbiguousContractDetails
End Enum

Public Enum HistDataWriteOptions
    HistDataWriteOptionWriteTradeBars = 1
    HistDataWriteOptionWriteBidAndAskBars = 2
    HistDataWriteOptionUpdateLatestBar = 4
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type BarWriterPendingComment
    Comment             As String
    Timestamp           As Date
End Type

Public Type BarWriterPendingBar
    Bar                 As Bar
    BarType             As BarTypes
    TimePeriod          As TimePeriod
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "HistDataUtils"

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

Public Function BarTypeToString(ByVal pBarType As BarTypes) As String
BarTypeToString = gBarTypeToString(pBarType)
End Function

Public Function CreateBarDataSpecifierFuture( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarTimePeriod As TimePeriod, _
                Optional ByVal pFromTime As Date, _
                Optional ByVal pToTime As Date, _
                Optional ByVal pMaxNumberOfBars As Long, _
                Optional ByVal pBarType As BarTypes, _
                Optional ByVal pClockFuture As IFuture, _
                Optional ByVal pExcludeCurrentBar As Boolean, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean, _
                Optional ByVal pCustomSessionStartTime As Date, _
                Optional ByVal pCustomSessionEndTime As Date) As IFuture
Const ProcName As String = "CreateBarDataSpecifierFuture"
On Error GoTo Err

Set CreateBarDataSpecifierFuture = gCreateBarDataSpecifierFuture(pContractFuture, pBarTimePeriod, pToTime, pFromTime, pMaxNumberOfBars, pBarType, pClockFuture, pExcludeCurrentBar, pIncludeBarsOutsideSession, pCustomSessionStartTime, pCustomSessionEndTime)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBufferedBarWriter( _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pContractFuture As IFuture) As IBarWriter
Const ProcName As String = "CreateBufferedBarWriter"
On Error GoTo Err

Set CreateBufferedBarWriter = gCreateBufferedBarWriter(pHistDataStore, pOutputMonitor, pContractFuture)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function RecordHistoricalBars( _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pOptions As HistDataWriteOptions, _
                Optional ByVal pSaveIntervalSeconds As Long = 5, _
                Optional ByVal pOutputMonitor As IBarOutputMonitor) As HistDataWriter
Const ProcName As String = "RecordHistoricalBars"
On Error GoTo Err

Dim lWriter As New HistDataWriter
lWriter.Initialise pContractFuture, pClockFuture, pStudyBase, pHistDataStore, pOutputMonitor, pOptions, pSaveIntervalSeconds
Set RecordHistoricalBars = lWriter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




