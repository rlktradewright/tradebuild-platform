Attribute VB_Name = "GBarFetcher"
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

Public Type BarRequestDetails
    BarTimePeriod       As TimePeriod
    FromDate            As Date
    ToDate              As Date
    NumberOfBars        As Long
    SessionTimes        As SessionTimes
    StartAtFromDate     As Boolean
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GBarFetcher"

Private Const MaxBarsToFetch                        As Long = 150000

Private Const TradingDaysPerYear                    As Double = 250
Private Const TradingDaysPerMonth                   As Double = 21
Private Const TradingDaysPerWeek                    As Double = 5

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

Public Function gCalcSessionTimes(ByVal pSpecifier As BarDataSpecifier, ByVal pInstrument As instrument) As SessionTimes
If Not pSpecifier.includeBarsOutsideSession Then
    gCalcSessionTimes.StartTime = IIf(pSpecifier.customSessionStartTime <> 0, pSpecifier.customSessionStartTime, pInstrument.SessionStartTime)
    gCalcSessionTimes.EndTime = IIf(pSpecifier.customSessionEndTime <> 0, pSpecifier.customSessionEndTime, pInstrument.SessionEndTime)
End If
End Function

Public Function gGenerateBarRequestDetails( _
                ByRef pSessionTimes As SessionTimes, _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pFromDate As Date, _
                ByVal pToDate As Date, _
                ByVal pMaxNumberOfBars As Long) As BarRequestDetails
Const ProcName As String = "gGenerateBarRequestDetails"
On Error GoTo Err

Dim lReqDetails As BarRequestDetails
lReqDetails.SessionTimes = pSessionTimes

Dim lMaxPermittedBars As Long
lMaxPermittedBars = getMaxNumberOfBars(pBarTimePeriod)

Dim lNumberOfTargetBars As Long
If pMaxNumberOfBars = 0 Or pMaxNumberOfBars > lMaxPermittedBars Then
    lNumberOfTargetBars = lMaxPermittedBars
Else
    lNumberOfTargetBars = pMaxNumberOfBars
End If

If pToDate <> 0 Then
    lReqDetails.ToDate = pToDate
    
    If pFromDate <> 0 Then
        lReqDetails.FromDate = pFromDate
    Else
        ' calculate the earliest possible date for the required number of bars
        lReqDetails.FromDate = OffsetBarStartTime(IIf(lReqDetails.ToDate = MaxDateValue, Now, lReqDetails.ToDate), _
                                    pBarTimePeriod, _
                                    -lNumberOfTargetBars, _
                                    pSessionTimes.StartTime, _
                                    pSessionTimes.EndTime)
        ' now move back a few sessions to allow for holidays, gaps etc
        lReqDetails.FromDate = GetOffsetSessionTimes(lReqDetails.FromDate, -gGetFetchOffset(pBarTimePeriod), pSessionTimes.StartTime, pSessionTimes.EndTime).StartTime
    End If
Else
    ' because pToDate is not supplied, we want the returned bars to start from pFromDate
    lReqDetails.StartAtFromDate = True
    If pFromDate <> 0 Then
        lReqDetails.FromDate = pFromDate
        ' calculate the latest possible date for the required number of bars
        lReqDetails.ToDate = OffsetBarStartTime(lReqDetails.FromDate, _
                                pBarTimePeriod, _
                                lNumberOfTargetBars, _
                                pSessionTimes.StartTime, _
                                pSessionTimes.EndTime)
        ' now move forward a few sessions to allow for holidays, gaps etc
        lReqDetails.ToDate = GetOffsetSessionTimes(lReqDetails.ToDate, gGetFetchOffset(pBarTimePeriod), pSessionTimes.StartTime, pSessionTimes.EndTime).StartTime
    End If
End If

Dim lBarLengthMinutes As Long

Select Case pBarTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    lBarLengthMinutes = 60
    lReqDetails.NumberOfBars = lNumberOfTargetBars * 24 / pBarTimePeriod.Length
Case TimePeriodUnits.TimePeriodHour
    lBarLengthMinutes = 60
    lReqDetails.NumberOfBars = lNumberOfTargetBars * pBarTimePeriod.Length
Case TimePeriodUnits.TimePeriodMinute
    If (pBarTimePeriod.Length Mod 60) = 0 Then
        lBarLengthMinutes = 60
        lReqDetails.NumberOfBars = lNumberOfTargetBars * (pBarTimePeriod.Length / 60)
    ElseIf (pBarTimePeriod.Length Mod 15) = 0 Then
        lBarLengthMinutes = 15
        lReqDetails.NumberOfBars = lNumberOfTargetBars * (pBarTimePeriod.Length / 15)
    ElseIf (pBarTimePeriod.Length Mod 5) = 0 Then
        lBarLengthMinutes = 5
        lReqDetails.NumberOfBars = lNumberOfTargetBars * (pBarTimePeriod.Length / 5)
    Else
        lBarLengthMinutes = 1
        lReqDetails.NumberOfBars = lNumberOfTargetBars * pBarTimePeriod.Length
    End If
Case TimePeriodUnits.TimePeriodMonth
    lBarLengthMinutes = 60
    lReqDetails.NumberOfBars = lNumberOfTargetBars * 24 * 21 * pBarTimePeriod.Length
Case TimePeriodUnits.TimePeriodWeek
    lBarLengthMinutes = 60
    lReqDetails.NumberOfBars = lNumberOfTargetBars * 24 * 5 * pBarTimePeriod.Length
Case TimePeriodUnits.TimePeriodYear
    lBarLengthMinutes = 60
    lReqDetails.NumberOfBars = lNumberOfTargetBars * 24 * 260 * pBarTimePeriod.Length
End Select

If lReqDetails.NumberOfBars > MaxBarsToFetch Then lReqDetails.NumberOfBars = MaxBarsToFetch
If lReqDetails.NumberOfBars = 0 Then lReqDetails.NumberOfBars = &H7FFFFFFF

Set lReqDetails.BarTimePeriod = GetTimePeriod(lBarLengthMinutes, TimePeriodMinute)

gGenerateBarRequestDetails = lReqDetails

Exit Function

Err:
gHandleUnexpectedError "ProcName", ModuleName
End Function

Public Sub gGenerateNextTickDataRequest( _
                ByRef pSessionTimes As SessionTimes, _
                ByVal pAppending As Boolean, _
                ByVal pPrevFromDate As Date, _
                ByVal pPrevToDate As Date, _
                ByRef pFromDate As Date, _
                ByRef pToDate As Date)
Const ProcName As String = "gGenerateNextTickDataRequest"
On Error GoTo Err

If pAppending Then
    gGenerateTickRequestDetails pSessionTimes, pPrevToDate, 0, pFromDate, pToDate
Else
    gGenerateTickRequestDetails pSessionTimes, 0, pPrevFromDate, pFromDate, pToDate
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gGenerateTickRequestDetails( _
                ByRef pSessionTimes As SessionTimes, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date, _
                ByRef pFromDate As Date, _
                ByRef pToDate As Date)
Const ProcName As String = "gGenerateTickRequestDetails"
On Error GoTo Err

If pEndDate = MaxDate Then pEndDate = Now + 10 * OneSecond

If pStartDate <> 0 And pEndDate <> 0 Then
    pFromDate = GetSessionTimes(pStartDate, _
                                pSessionTimes.StartTime, _
                                pSessionTimes.EndTime).StartTime
    pToDate = pEndDate
End If

If pStartDate <> 0 And pEndDate = 0 Then
    pFromDate = GetSessionTimes(pStartDate, _
                                pSessionTimes.StartTime, _
                                pSessionTimes.EndTime).StartTime
    pToDate = GetOffsetSessionTimes(pFromDate, _
                                1, _
                                pSessionTimes.StartTime, _
                                pSessionTimes.EndTime).StartTime
End If

If pStartDate = 0 And pEndDate <> 0 Then
    pToDate = pEndDate
    pFromDate = GetOffsetSessionTimes(pToDate, _
                                -1, _
                                pSessionTimes.StartTime, _
                                pSessionTimes.EndTime).StartTime
End If

If pStartDate = 0 And pEndDate = 0 Then
    pToDate = Now
    pFromDate = GetOffsetSessionTimes(pToDate, _
                                -1, _
                                pSessionTimes.StartTime, _
                                pSessionTimes.EndTime).StartTime
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gGetFetchOffset(ByVal pTimePeriod As TimePeriod) As Long
Select Case pTimePeriod.Units
    Case TimePeriodMinute, _
            TimePeriodHour
            gGetFetchOffset = TradingDaysPerWeek
    Case TimePeriodDay, _
            TimePeriodWeek, _
            TimePeriodMonth, _
            TimePeriodYear
        gGetFetchOffset = TradingDaysPerMonth
End Select
End Function

Public Function gSetupFetchBarsCommand( _
                ByVal pInstrumentID As Long, _
                ByVal BarType As Long, _
                ByRef pReqDetails As BarRequestDetails) As Command
Const ProcName As String = "gSetupFetchBarsCommand"
On Error GoTo Err

Dim lCmd As New ADODB.Command
lCmd.CommandType = adCmdStoredProc

lCmd.CommandText = "FetchBarData"

Dim param As ADODB.Parameter
' @InstrumentID
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            pInstrumentID)
lCmd.Parameters.Append param

' @BarType
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            BarType)
lCmd.Parameters.Append param

' @BarLength
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            pReqDetails.BarTimePeriod.Length)
lCmd.Parameters.Append param

' @NumberRequired
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            pReqDetails.NumberOfBars)
lCmd.Parameters.Append param

Dim lFromTime As Date: lFromTime = gRoundTimeToSecond(pReqDetails.FromDate)
Dim lToTime As Date: lToTime = gRoundTimeToSecond(pReqDetails.ToDate)

If lFromTime < #1/1/1900# Then lFromTime = #1/1/1900# ' don't exceed range of SmallDateTime
If lToTime < #1/1/1900# Then lToTime = #1/1/1900# ' don't exceed range of SmallDateTime
' @From
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            lFromTime)
lCmd.Parameters.Append param

' @To
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            lToTime)
lCmd.Parameters.Append param

' @SessionStart
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTime, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            pReqDetails.SessionTimes.StartTime)
lCmd.Parameters.Append param

' @SessionEnd
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTime, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            pReqDetails.SessionTimes.EndTime)
lCmd.Parameters.Append param

' @StartAtFromDate
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            IIf(pReqDetails.StartAtFromDate, 1, 0))
lCmd.Parameters.Append param

' @Ascending
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            1)
lCmd.Parameters.Append param

Set gSetupFetchBarsCommand = lCmd

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gSetupFetchTicksCommand( _
                ByVal pInstrumentID As Long, _
                ByVal pFromTime As Date, _
                ByVal pToTime As Date, _
                ByRef sessTimes As SessionTimes) As Command
Const ProcName As String = "gSetupFetchTicksCommand"
On Error GoTo Err

pFromTime = gRoundTimeToSecond(pFromTime)
pToTime = gRoundTimeToSecond(pToTime)

Dim lCmd As New ADODB.Command
lCmd.CommandType = adCmdStoredProc

lCmd.CommandText = "FetchTickData"

Dim param As ADODB.Parameter

' @InstrumentID
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adInteger, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            pInstrumentID)
lCmd.Parameters.Append param

' @From
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            pFromTime)
lCmd.Parameters.Append param

' @To
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTimeStamp, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            pToTime)
lCmd.Parameters.Append param

' @SessionStart
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTime, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            sessTimes.StartTime)
lCmd.Parameters.Append param

' @SessionEnd
Set param = lCmd.CreateParameter(, _
                            DataTypeEnum.adDBTime, _
                            ParameterDirectionEnum.adParamInput, _
                            , _
                            sessTimes.EndTime)
lCmd.Parameters.Append param

Set gSetupFetchTicksCommand = lCmd

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gUseTickData(ByVal pUnits As TimePeriodUnits) As Boolean
gUseTickData = (pUnits = TimePeriodVolume Or _
    pUnits = TimePeriodTickMovement Or _
    pUnits = TimePeriodTickVolume Or _
    pUnits = TimePeriodSecond)
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getMaxNumberOfBars(ByVal pTimePeriod As TimePeriod) As Long
Dim lMaxBars As Long
Select Case pTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    lMaxBars = MaxBarsToFetch / 24 / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodHour
    lMaxBars = MaxBarsToFetch / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodMinute
    lMaxBars = MaxBarsToFetch / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodMonth
    lMaxBars = MaxBarsToFetch / TradingDaysPerMonth / 24 / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodWeek
    lMaxBars = MaxBarsToFetch / TradingDaysPerWeek / 24 / pTimePeriod.Length
Case TimePeriodUnits.TimePeriodYear
    lMaxBars = MaxBarsToFetch / TradingDaysPerYear / 24 / pTimePeriod.Length
End Select
getMaxNumberOfBars = lMaxBars
End Function




