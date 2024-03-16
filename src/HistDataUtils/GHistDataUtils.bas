Attribute VB_Name = "GHistDataUtils"
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

Private Const ModuleName                            As String = "GHistDataUtils"

Public Const Latest                                 As String = "LATEST"
Public Const Today                                  As String = "TODAY"
Public Const Tomorrow                               As String = "TOMORROW"
Public Const Yesterday                              As String = "YESTERDAY"
Public Const EndOfWeek                              As String = "ENDOFWEEK"
Public Const StartOfWeek                            As String = "STARTOFWEEK"
Public Const StartOfPreviousWeek                    As String = "STARTOFPREVIOUSWEEK"

Public Const All                                    As Long = &H7FFFFFFF

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
Select Case pBarType
Case BarTypeTrade
    BarTypeToString = "TRADE"
Case BarTypeBid
    BarTypeToString = "BID"
Case BarTypeAsk
    BarTypeToString = "ASK"
Case Else
    AssertArgument False, "Invalid bar type"
End Select
End Function

Public Function CreateBarDataSpecifier( _
                ByVal pBarTimePeriod As TimePeriod, _
                Optional ByVal pFromTime As Date, _
                Optional ByVal pToTime As Date, _
                Optional ByVal pMaxNumberOfBars As Long, _
                Optional ByVal pBarType As BarTypes, _
                Optional ByVal pExcludeCurrentbar As Boolean, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean, _
                Optional ByVal pNormaliseDailyTimestamps As Boolean, _
                Optional ByVal pCustomSessionStartTime As Date, _
                Optional ByVal pCustomSessionEndTime As Date) As BarDataSpecifier
Const ProcName As String = "CreateBarDataSpecifier"
On Error GoTo Err

AssertArgument Not pBarTimePeriod Is Nothing

Dim lBarDataSpecifier As New BarDataSpecifier
lBarDataSpecifier.Initialise _
                            pBarTimePeriod, _
                            pToTime, _
                            pFromTime, _
                            pMaxNumberOfBars, _
                            pBarType, _
                            pExcludeCurrentbar, _
                            pIncludeBarsOutsideSession, _
                            pNormaliseDailyTimestamps, _
                            pCustomSessionStartTime, _
                            pCustomSessionEndTime
Set CreateBarDataSpecifier = lBarDataSpecifier

Exit Function

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBufferedBarWriter( _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pOutputMonitor As IBarOutputMonitor, _
                ByVal pContractFuture As IFuture) As IBarWriter
Const ProcName As String = "CreateBufferedBarWriter"
On Error GoTo Err

Dim lBufferedWriter As New BufferedBarWriter
Dim lWriter As IBarWriter
Set lWriter = pHistDataStore.CreateBarWriter(lBufferedWriter, pContractFuture)
lBufferedWriter.Initialise pOutputMonitor, lWriter, pContractFuture
Set CreateBufferedBarWriter = lBufferedWriter

Exit Function

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
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

GHistData.Logger.Log "RecordHistoricalBars", ProcName, ModuleName, LogLevelHighDetail

Dim lWriter As New HistDataWriter
lWriter.Initialise pContractFuture, pClockFuture, pStudyBase, pHistDataStore, pOutputMonitor, pOptions, pSaveIntervalSeconds
Set RecordHistoricalBars = lWriter

Exit Function

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function SpecialTimeToDate( _
                ByVal pSpecialTime As String, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                Optional ByVal pClock As Clock) As Date
Const ProcName As String = "SpecialTimeToDate"
On Error GoTo Err

Dim lTimestamp As Date
If pClock Is Nothing Then
    lTimestamp = GetTimestamp
Else
    lTimestamp = pClock.Timestamp
End If

pSpecialTime = UCase$(pSpecialTime)

If pSpecialTime = Today Then
    SpecialTimeToDate = todayDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = Yesterday Then
    SpecialTimeToDate = yesterdayDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = StartOfWeek Then
    SpecialTimeToDate = startOfWeekDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = StartOfPreviousWeek Then
    SpecialTimeToDate = startOfPreviousWeekDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = Latest Then
    SpecialTimeToDate = MaxDate
ElseIf pSpecialTime = Tomorrow Then
    SpecialTimeToDate = tomorrowDate(lTimestamp, pSessionStartTime, pSessionEndTime)
ElseIf pSpecialTime = EndOfWeek Then
    SpecialTimeToDate = endOfWeekDate(lTimestamp, pSessionStartTime, pSessionEndTime)
Else
    AssertArgument False, "Invalid special time '" & pSpecialTime & "'"
End If

Exit Function

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function endOfWeekDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
endOfWeekDate = GetSessionTimes( _
                 Int(pTimestamp) - DatePart("w", pTimestamp, vbMonday) + vbFriday - 1 + 0.5, _
                 pSessionStartTime, _
                 pSessionEndTime).StartTime
End Function

Private Function startOfPreviousWeekDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
startOfPreviousWeekDate = GetSessionTimes( _
                Int(pTimestamp) - DatePart("w", pTimestamp, vbMonday) + vbMonday - 8 + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function

Private Function startOfWeekDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
startOfWeekDate = GetSessionTimes( _
                Int(pTimestamp) - DatePart("w", pTimestamp, vbMonday) + vbMonday - 1 + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function

Private Function todayDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
todayDate = GetSessionTimes( _
                Int(WorkingDayDate(WorkingDayNumber(pTimestamp), pTimestamp)) + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function

Private Function tomorrowDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
tomorrowDate = GetSessionTimes( _
                Int(WorkingDayDate(WorkingDayNumber(pTimestamp) + 1, pTimestamp)) + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function

Private Function yesterdayDate( _
                ByVal pTimestamp As Date, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
yesterdayDate = GetSessionTimes( _
                Int(WorkingDayDate(WorkingDayNumber(pTimestamp) - 1, pTimestamp)) + 0.5, _
                pSessionStartTime, _
                pSessionEndTime).StartTime
End Function






