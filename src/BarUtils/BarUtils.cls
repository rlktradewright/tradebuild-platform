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

BarEndTime = gBarEndTime(Timestamp, _
                            BarTimePeriod, _
                            gNormaliseTime(SessionStartTime), _
                            gNormaliseTime(SessionEndTime))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BarLength( _
                ByVal BarTimePeriod As TimePeriod) As Date
Const ProcName As String = "BarLength"

On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case Else
    AssertArgument False, "Can't calculate Bar length for this time unit"
End Select
BarLength = gCalcBarLength(BarTimePeriod)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal SessionStartTime As Date) As Date
Const ProcName As String = "BarStartTime"

On Error GoTo Err

BarStartTime = gBarStartTime(Timestamp, _
                            BarTimePeriod, _
                            gNormaliseTime(SessionStartTime))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Set CreateBar = New Bar
CreateBar.Initialise Timestamp, _
                    OpenValue, _
                    HighValue, _
                    LowValue, _
                    CloseValue, _
                    Volume, _
                    TickVolume, _
                    OpenInterest

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilder( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pSession As Session, _
                Optional ByVal pTickSize As Double, _
                Optional ByVal pNumberOfBarsToCache As Long, _
                Optional ByVal pSave As Boolean = True) As BarsBuilder
Const ProcName As String = "CreateBarsBuilder"
On Error GoTo Err

AssertArgument Not pBarTimePeriod Is Nothing, "pBarTimePeriod is Nothing"
AssertArgument Not pSession Is Nothing, "pSession is Nothing"

Dim lBarsBuilder As New BarsBuilder
lBarsBuilder.Initialise pBarTimePeriod, pSession, pTickSize, pNumberOfBarsToCache, pSave
Set CreateBarsBuilder = lBarsBuilder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderFuture( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pSessionFuture As IFuture, _
                Optional ByVal pTickSize As Double, _
                Optional ByVal pNumberOfBarsToCache As Long, _
                Optional ByVal pSave As Boolean = True) As IFuture
Const ProcName As String = "CreateBarsBuilderFuture"
On Error GoTo Err

AssertArgument Not pBarTimePeriod Is Nothing, "pBarTimePeriod is Nothing"
AssertArgument Not pSessionFuture Is Nothing, "pSessionFuture is Nothing"

Dim lFutureBuilder As New BarsBuilderFutureBuilder
lFutureBuilder.Initialise pBarTimePeriod, pSessionFuture, pTickSize, pNumberOfBarsToCache, pSave
Set CreateBarsBuilderFuture = lFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderWithInitialBars( _
                ByVal pBars As Bars, _
                ByVal pSession As Session, _
                Optional ByVal pTickSize As Double) As IFuture
Const ProcName As String = "CreateBarsBuilderWithInitialBars"
On Error GoTo Err

AssertArgument Not pBars Is Nothing, "pBars Is Nothing"
AssertArgument Not pSession Is Nothing, "pSession is Nothing"
    
Dim lBarsBuilder As New BarsBuilder
lBarsBuilder.InitialiseWithInitialBars pBars, pSession, pTickSize
Set CreateBarsBuilderWithInitialBars = lBarsBuilder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarsBuilderWithInitialBarsFuture( _
                ByVal pBarsFuture As IFuture, _
                ByVal pSessionFuture As IFuture, _
                Optional ByVal pTickSize As Double) As IFuture
Const ProcName As String = "CreateBarsBuilderWithInitialBarsFuture"
On Error GoTo Err

AssertArgument Not pBarsFuture Is Nothing, "pBarsFuture Is Nothing"
AssertArgument Not pSessionFuture Is Nothing, "pSessionFuture is Nothing"
    
Dim lFutureBuilder As New BarsBuilderFutureBuilder
lFutureBuilder.InitialiseWithInitialBars pBarsFuture, pSessionFuture, pTickSize
Set CreateBarsBuilderWithInitialBarsFuture = lFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateValueCache( _
                ByVal CyclicSize As Long, _
                ByVal ValueName As String) As ValueCache
Const ProcName As String = "CreateValueCache"
On Error GoTo Err

Set CreateValueCache = New ValueCache
CreateValueCache.Initialise CyclicSize, ValueName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateVolumeParser( _
                ByVal pSessionFuture As IFuture) As VolumeParser
Const ProcName As String = "CreateVolumeParser"
On Error GoTo Err

Set CreateVolumeParser = New VolumeParser
CreateVolumeParser.Initialise pSessionFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MaxNumberOfBarsInTimespan( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pStartTime As Date, _
                Optional ByVal pEndTime As Date, _
                Optional ByVal pSessionStartTime As Date, _
                Optional ByVal pSessionEndTime As Date) As Long
Const ProcName As String = "MaxNumberOfBarsInTimespan"
On Error GoTo Err

AssertArgument pStartTime <> 0, "pStartTime must be supplied"

Select Case pBarTimePeriod.Units
    Case TimePeriodNone, TimePeriodTickMovement, TimePeriodTickVolume, TimePeriodVolume
        AssertArgument False, "Must be a fixed time period"
End Select
        
If pEndTime = 0 Then pEndTime = Now

Dim lStartTime As Date
lStartTime = BarStartTime(pStartTime, pBarTimePeriod, pSessionStartTime)

Dim lEndTime As Date
lEndTime = OffsetBarStartTime(pEndTime, pBarTimePeriod, 1, pSessionStartTime, pSessionEndTime)

Select Case pBarTimePeriod.Units
    Case TimePeriodSecond
        MaxNumberOfBarsInTimespan = Int((86400# * (lEndTime - lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodMinute
        MaxNumberOfBarsInTimespan = Int((1440# * (lEndTime - lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodHour
        MaxNumberOfBarsInTimespan = Int((24# * (lEndTime - lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodDay
        MaxNumberOfBarsInTimespan = Int(((lEndTime - lStartTime) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodWeek
        MaxNumberOfBarsInTimespan = Int(((lEndTime - lStartTime) / 7 + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodMonth
        MaxNumberOfBarsInTimespan = Int(((Month(lEndTime) - Month(lStartTime)) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
    Case TimePeriodYear
        MaxNumberOfBarsInTimespan = Int(((Year(lEndTime) - Year(lStartTime)) + pBarTimePeriod.Length - 1) / pBarTimePeriod.Length)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long
Const ProcName As String = "NumberOfBarsInSession"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case Else
    AssertArgument False, "Can't calculate number of Bars in session for this time unit"
End Select

NumberOfBarsInSession = gCalcNumberOfBarsInSession( _
                                                BarTimePeriod, _
                                                gNormaliseTime(SessionStartTime), _
                                                gNormaliseTime(SessionEndTime))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "OffsetBarStartTime"
On Error GoTo Err

AssertArgument BarTimePeriod.Units <> TimePeriodNone, "Invalid time Units argument"

OffsetBarStartTime = gCalcOffsetBarStartTime( _
                                                Timestamp, _
                                                BarTimePeriod, _
                                                offset, _
                                                gNormaliseTime(SessionStartTime), _
                                                gNormaliseTime(SessionEndTime))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


