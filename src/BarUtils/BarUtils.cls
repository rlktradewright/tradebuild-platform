VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "BarUtils"

'@================================================================================
' Enums
'@================================================================================

Public Enum BarChangeTypes
    BarPriceChanged = 1
    BarVolumeChanged = 2
    BarOpenInterestChanged = 4
    BarTickVolumeChanged = 8 ' only fired when TickVolume is specifically incremented
                            ' BarPriceChanged implicitly includes BarTickVolumeChanged
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type SValue
    Value           As Variant
    BarNumber       As Long
    Timestamp       As Date
    BarStartTime    As Date
End Type

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

Public Function BarEndTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "BarEndTime"
On Error GoTo Err

BarEndTime = GBarUtils.BarEndTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function BarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "BarStartTime"
On Error GoTo Err

BarStartTime = GBarUtils.BarStartTime(Timestamp, BarTimePeriod, SessionStartTime, SessionEndTime)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBar( _
                ByVal Timestamp As Date, _
                ByVal OpenValue As Double, _
                ByVal HighValue As Double, _
                ByVal LowValue As Double, _
                ByVal CloseValue As Double, _
                Optional ByVal Volume As Long, _
                Optional ByVal TickVolume As Long, _
                Optional ByVal OpenInterest As Long) As Bar
Const ProcName As String = "CreateBar"
On Error GoTo Err

Set CreateBar = GBarUtils.CreateBar(Timestamp, _
                    OpenValue, _
                    HighValue, _
                    LowValue, _
                    CloseValue, _
                    Volume, _
                    TickVolume, _
                    OpenInterest)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilder( _
                ByVal pBarTimePeriod As TimePeriod, _
                Optional ByVal pSession As Session, _
                Optional ByVal pTickSize As Double, _
                Optional ByVal pNumberOfBarsToCache As Long, _
                Optional ByVal pNormaliseDailyTimestamps As Boolean, _
                Optional ByVal pSave As Boolean = True) As BarsBuilder
Const ProcName As String = "CreateBarsBuilder"
On Error GoTo Err

Set CreateBarsBuilder = GBarUtils.CreateBarsBuilder(pBarTimePeriod, pSession, pTickSize, pNumberOfBarsToCache, pNormaliseDailyTimestamps, pSave)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderFuture( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pSessionFuture As IFuture, _
                Optional ByVal pTickSize As Double, _
                Optional ByVal pNumberOfBarsToCache As Long, _
                Optional ByVal pNormaliseDailyTimestamps As Boolean, _
                Optional ByVal pSave As Boolean = True) As IFuture
Const ProcName As String = "CreateBarsBuilderFuture"
On Error GoTo Err

Set CreateBarsBuilderFuture = GBarUtils.CreateBarsBuilderFuture(pBarTimePeriod, pSessionFuture, pTickSize, pNumberOfBarsToCache, pNormaliseDailyTimestamps, pSave)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderWithInitialBars( _
                ByVal pBars As Bars, _
                ByVal pSession As Session, _
                Optional ByVal pTickSize As Double) As BarsBuilder
Const ProcName As String = "CreateBarsBuilderWithInitialBars"
On Error GoTo Err

Set CreateBarsBuilderWithInitialBars = GBarUtils.CreateBarsBuilderWithInitialBars(pBars, pSession, pTickSize)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderWithInitialBarsFuture( _
                ByVal pBarsFuture As IFuture, _
                ByVal pSession As Session, _
                Optional ByVal pTickSize As Double) As IFuture
Const ProcName As String = "CreateBarsBuilderWithInitialBarsFuture"
On Error GoTo Err

Set CreateBarsBuilderWithInitialBarsFuture = GBarUtils.CreateBarsBuilderWithInitialBarsFuture(pBarsFuture, pSession, pTickSize)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateValueCache( _
                ByVal CyclicSize As Long, _
                ByVal ValueName As String) As ValueCache
Const ProcName As String = "CreateValueCache"
On Error GoTo Err

Set CreateValueCache = GBarUtils.CreateValueCache(CyclicSize, ValueName)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateVolumeParser( _
                ByVal pSession As Session) As VolumeParser
Const ProcName As String = "CreateVolumeParser"
On Error GoTo Err

Set CreateVolumeParser = GBarUtils.CreateVolumeParser(pSession)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub GetTimespanData( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByRef pFromTime As Date, _
                ByRef pToTime As Date, _
                ByRef pFromSessionTimes As SessionTimes, _
                ByRef pToSessionTimes As SessionTimes, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date)
Const ProcName As String = "GetTimespanData"
On Error GoTo Err

GBarUtils.GetTimespanData pBarTimePeriod, _
                pFromTime, _
                pToTime, _
                pFromSessionTimes, _
                pToSessionTimes, _
                pSessionStartTime, _
                pSessionEndTime

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub


Public Function MaxNumberOfBarsInTimespan( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pStartTime As Date, _
                Optional ByVal pEndTime As Date, _
                Optional ByVal pSessionStartTime As Date, _
                Optional ByVal pSessionEndTime As Date) As Long
Const ProcName As String = "MaxNumberOfBarsInTimespan"
On Error GoTo Err

MaxNumberOfBarsInTimespan = GBarUtils.MaxNumberOfBarsInTimespan( _
                                                        pBarTimePeriod, _
                                                        pStartTime, _
                                                        pEndTime, _
                                                        pSessionStartTime, _
                                                        pSessionEndTime)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function MaxNumberOfBarsInTimespanNormalized( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date, _
                ByRef pStartSessionTimes As SessionTimes, _
                ByRef pEndSessionTimes As SessionTimes) As Long
Const ProcName As String = "MaxNumberOfBarsInTimespanNormalized"
On Error GoTo Err

MaxNumberOfBarsInTimespanNormalized = GBarUtils.MaxNumberOfBarsInTimespanNormalized( _
                                            pBarTimePeriod, _
                                            pStartTime, _
                                            pEndTime, _
                                            pStartSessionTimes, _
                                            pEndSessionTimes)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function


Public Function NormaliseBarTimestamp( _
                ByVal pTimestamp As Date, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date) As Date
NormaliseBarTimestamp = GBarUtils.NormaliseTimestamp(pTimestamp, _
                                            pTimePeriod, _
                                            pSessionStartTime, _
                                            pSessionEndTime)
End Function

Public Function NumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long
Const ProcName As String = "NumberOfBarsInSession"
On Error GoTo Err

NumberOfBarsInSession = GBarUtils.NumberOfBarsInSession( _
                                                BarTimePeriod, _
                                                SessionStartTime, _
                                                SessionEndTime)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "OffsetBarStartTime"
On Error GoTo Err

OffsetBarStartTime = GBarUtils.OffsetBarStartTime( _
                                Timestamp, _
                                BarTimePeriod, _
                                offset, _
                                SessionStartTime, _
                                SessionEndTime)

Exit Function

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


