VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeframeUtils"
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

Private Const ProjectName               As String = "TimeframeUtils26"
Private Const ModuleName                As String = "GlobalsClass"

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

Public Enum BarTypes
    BarTypeTrade = 0
    BarTypeBid = 1
    BarTypeAsk = 2
End Enum

Public Enum SessionChangeTypes
    SessionChangeNone
    SessionChangeEnd
    SessionChangeStart
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type SessionEventData
    Source              As Object
    Timestamp           As Date
    ChangeType          As SessionChangeTypes
End Type

Public Type SessionTimes
    StartTime       As Date
    EndTime         As Date
End Type

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
Dim failpoint As String
On Error GoTo Err

BarEndTime = gBarEndTime(Timestamp, _
                            BarTimePeriod, _
                            gNormaliseTime(SessionStartTime), _
                            gNormaliseTime(SessionEndTime))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function BarLength( _
                ByVal BarTimePeriod As TimePeriod) As Date
Const ProcName As String = "BarLength"
Dim failpoint As String
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Can't calculate Bar length for this time unit"
End Select
BarLength = gCalcBarLength(BarTimePeriod)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function BarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal SessionStartTime As Date) As Date
Const ProcName As String = "BarStartTime"
Dim failpoint As String
On Error GoTo Err

BarStartTime = gBarStartTime(Timestamp, _
                            BarTimePeriod, _
                            gNormaliseTime(SessionStartTime))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
Dim failpoint As String
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function CreateBarsBuilder( _
                ByVal BarTimePeriod As TimePeriod, _
                Optional ByVal pSession As Session, _
                Optional ByVal TickSize As Double, _
                Optional ByVal numberOfBarsToCache As Long) As BarsBuilder
Set CreateBarsBuilder = New BarsBuilder
If pSession Is Nothing Then
    Dim lSB As New SessionBuilder
    Set pSession = lSB.Session
End If
CreateBarsBuilder.Initialise BarTimePeriod, _
                            pSession, TickSize, numberOfBarsToCache
End Function

Public Function CreateValueCache( _
                ByVal CyclicSize As Long, _
                ByVal ValueName As String) As ValueCache
Const ProcName As String = "CreateValueCache"
Dim failpoint As String
On Error GoTo Err

Set CreateValueCache = New ValueCache
CreateValueCache.Initialise CyclicSize, ValueName

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function CreateVolumeParser( _
                ByVal pSession As Session) As VolumeParser
Const ProcName As String = "CreateVolumeParser"
Dim failpoint As String
On Error GoTo Err

Set CreateVolumeParser = New VolumeParser
CreateVolumeParser.Initialise pSession

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetSessionTimes( _
                ByVal Timestamp As Date, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date) As SessionTimes
Const ProcName As String = "GetSessionTimes"
Dim failpoint As String
On Error GoTo Err

GetSessionTimes = gCalcSessionTimes(Timestamp, _
                            gNormaliseTime(StartTime), _
                            gNormaliseTime(EndTime))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetOffsetSessionTimes( _
                ByVal Timestamp As Date, _
                ByVal offset As Long, _
                Optional ByVal StartTime As Date, _
                Optional ByVal EndTime As Date) As SessionTimes
Const ProcName As String = "GetOffsetSessionTimes"
Dim failpoint As String
On Error GoTo Err

GetOffsetSessionTimes = gCalcOffsetSessionTimes(Timestamp, _
                            offset, _
                            gNormaliseTime(StartTime), _
                            gNormaliseTime(EndTime))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function NumberOfBarsInSession( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal SessionStartTime As Date, _
                ByVal SessionEndTime As Date) As Long
Const ProcName As String = "NumberOfBarsInSession"
Dim failpoint As String
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Can't calculate number of Bars in session for this time unit"
End Select

NumberOfBarsInSession = gCalcNumberOfBarsInSession( _
                                                BarTimePeriod, _
                                                gNormaliseTime(SessionStartTime), _
                                                gNormaliseTime(SessionEndTime))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function OffsetBarStartTime( _
                ByVal Timestamp As Date, _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal offset As Long, _
                Optional ByVal SessionStartTime As Date, _
                Optional ByVal SessionEndTime As Date) As Date
Const ProcName As String = "OffsetBarStartTime"
Dim failpoint As String
On Error GoTo Err

If BarTimePeriod.Units = TimePeriodNone Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid time Units argument"
End If
OffsetBarStartTime = gCalcOffsetBarStartTime( _
                                                Timestamp, _
                                                BarTimePeriod, _
                                                offset, _
                                                gNormaliseTime(SessionStartTime), _
                                                gNormaliseTime(SessionEndTime))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================


