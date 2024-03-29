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

Public Property Get AllBars() As Long
AllBars = All
End Property

Public Property Get DateLatest() As String
DateLatest = Latest
End Property

Public Property Get DateToday() As String
DateToday = Today
End Property

Public Property Get DateTomorrow() As String
DateTomorrow = Tomorrow
End Property

Public Property Get DateYesterday() As String
DateYesterday = Yesterday
End Property

Public Property Get DateEndOfWeek() As String
DateEndOfWeek = EndOfWeek
End Property

Public Property Get DateStartOfWeek() As String
DateStartOfWeek = StartOfWeek
End Property

Public Property Get DateStartOfPreviousWeek() As String
DateStartOfPreviousWeek = StartOfPreviousWeek
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function BarTypeToString(ByVal pBarType As BarTypes) As String
BarTypeToString = gBarTypeToString(pBarType)
End Function

Public Function CreateBarDataSpecifier( _
                ByVal pBarTimePeriod As TimePeriod, _
                Optional ByVal pFromTime As Date, _
                Optional ByVal pToTime As Date, _
                Optional ByVal pMaxNumberOfBars As Long, _
                Optional ByVal pBarType As BarTypes, _
                Optional ByVal pExcludeCurrentBar As Boolean, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean, _
                Optional ByVal pNormaliseDailyTimestamps As Boolean, _
                Optional ByVal pCustomSessionStartTime As Date, _
                Optional ByVal pCustomSessionEndTime As Date) As BarDataSpecifier
Const ProcName As String = "CreateBarDataSpecifier"
On Error GoTo Err

Set CreateBarDataSpecifier = gCreateBarDataSpecifier( _
                                        pBarTimePeriod, _
                                        pToTime, _
                                        pFromTime, _
                                        pMaxNumberOfBars, _
                                        pBarType, _
                                        pExcludeCurrentBar, _
                                        pIncludeBarsOutsideSession, _
                                        pNormaliseDailyTimestamps, _
                                        pCustomSessionStartTime, _
                                        pCustomSessionEndTime)

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

gLogger.Log "RecordHistoricalBars", ProcName, ModuleName, LogLevelHighDetail

Dim lWriter As New HistDataWriter
lWriter.Initialise pContractFuture, pClockFuture, pStudyBase, pHistDataStore, pOutputMonitor, pOptions, pSaveIntervalSeconds
Set RecordHistoricalBars = lWriter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SpecialTimeToDate( _
                ByVal pSpecialTime As String, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                Optional ByVal pClock As Clock) As Date
Const ProcName As String = "SpecialTimeToDate"
On Error GoTo Err

SpecialTimeToDate = gSpecialTimeToDate( _
                                pSpecialTime, _
                                pSessionStartTime, _
                                pSessionEndTime, _
                                pClock)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




