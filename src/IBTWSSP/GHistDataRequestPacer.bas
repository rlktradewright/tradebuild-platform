Attribute VB_Name = "GHistDataRequests"
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

Private Const ModuleName                            As String = "GHistDataRequestPacer"

Private Const MaxTwsHistRequestsPerPeriod           As Long = 60

Private Const TwsHistRequestPeriod                  As Double = 10# / (60# * 24#)

Private Const TwsHistRequestRepeatInterval          As Double = 15# / (60# * 60# * 24#)

Private Const TwsHistRequestContractAndTickTypePeriod   As Double = 9# / (60# * 60# * 24#)
Private Const TwsHistRequestContractAndTickTypeRequestLimit   As Long = 6

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

Public Property Get gDelayTillNextSubmission() As Long
Const ProcName As String = "gDelayTillNextSubmission"
On Error GoTo Err

Dim lTimestamp As Date
Dim lEarliestSubmissionTime As Date

lTimestamp = GetTimestamp

If submittedHistDataRequests.Count = submittedHistDataRequests.CyclicSize Then
    lEarliestSubmissionTime = submittedHistDataRequests.GetSValue(1).Timestamp + TwsHistRequestPeriod
    If lTimestamp < lEarliestSubmissionTime Then
        gDelayTillNextSubmission = ((lEarliestSubmissionTime - lTimestamp) * 86400 + 5) * 1000
    Else
        gDelayTillNextSubmission = 0
    End If
Else
    gDelayTillNextSubmission = 0
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gGetEarliestSubmissionTime(ByRef pRequest As HistoricalDataRequest) As Date
Const ProcName As String = "gGetEarliestSubmissionTime"
On Error GoTo Err

gGetEarliestSubmissionTime = getEarliestTimeForIdenticalRequest(pRequest)

Dim lContractTime As Date
lContractTime = getEarliestTimeForContractAndTickType(pRequest)

If lContractTime > gGetEarliestSubmissionTime Then gGetEarliestSubmissionTime = lContractTime

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gRecordSubmission(ByRef pRequest As HistoricalDataRequest) As Date
Const ProcName As String = "gRecordSubmission"
On Error GoTo Err

Dim lTimestamp As Date
Dim lKey As String

lTimestamp = GetTimestamp

submittedHistDataRequests.AddValue pRequest, 0, lTimestamp, 0

lKey = historicalDataRequestToString(pRequest)
On Error Resume Next
requestTimes.Remove lKey
On Error GoTo Err

requestTimes.Add pRequest, lKey

Dim lCache As ValueCache

Set lCache = getContractAndTickTypeCache(pRequest)
lCache.AddValue pRequest, 0, lTimestamp, 0

gRecordSubmission = lTimestamp

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function contractAndTickTypeToString(ByRef pRequest As HistoricalDataRequest) As String
Const ProcName As String = "contractAndTickTypeToString"
On Error GoTo Err

contractAndTickTypeToString = pRequest.Contract.specifier.key & vbNullString & _
                                pRequest.WhatToShow

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getContractAndTickTypeCache(ByRef pRequest As HistoricalDataRequest) As ValueCache
Const ProcName As String = "getContractAndTickTypeCache"
On Error GoTo Err

Dim lCache As ValueCache
Dim lKey As String

lKey = contractAndTickTypeToString(pRequest)

On Error Resume Next
Set lCache = requestTimesForContractAndTickType(lKey)
On Error GoTo Err

If lCache Is Nothing Then
    Set lCache = CreateValueCache(TwsHistRequestContractAndTickTypeRequestLimit, "Time")
    requestTimesForContractAndTickType.Add lCache, lKey
End If

Set getContractAndTickTypeCache = lCache

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getEarliestTimeForContractAndTickType(ByRef pRequest As HistoricalDataRequest) As Date
Const ProcName As String = "getEarliestTimeForContractAndTickType"
On Error GoTo Err

Dim lCache As ValueCache

Set lCache = getContractAndTickTypeCache(pRequest)

If lCache.Count = lCache.CyclicSize Then
    getEarliestTimeForContractAndTickType = lCache.GetSValue(1).Timestamp + TwsHistRequestContractAndTickTypePeriod
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getEarliestTimeForIdenticalRequest(ByRef pRequest As HistoricalDataRequest) As Date
Const ProcName As String = "getEarliestTimeForIdenticalRequest"
On Error GoTo Err

Dim lLastTime As Date

On Error Resume Next
lLastTime = CDate(requestTimes(historicalDataRequestToString(pRequest)))
If lLastTime <> 0 Then getEarliestTimeForIdenticalRequest = lLastTime + TwsHistRequestRepeatInterval

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function historicalDataRequestToString(ByRef pRequest As HistoricalDataRequest) As String
Const ProcName As String = "historicalDataRequestToString"
On Error GoTo Err

historicalDataRequestToString = pRequest.BarSizeSetting & vbNullString & _
                                pRequest.Contract.specifier.key & vbNullString & _
                                pRequest.Duration & vbNullString & _
                                pRequest.EndDateTime & vbNullString & _
                                pRequest.WhatToShow

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function requestTimes() As Collection
Const ProcName As String = "requestTimes"
On Error GoTo Err

Static sRequestTimes As Collection

If sRequestTimes Is Nothing Then Set sRequestTimes = New Collection
Set requestTimes = sRequestTimes

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function requestTimesForContractAndTickType() As Collection
Const ProcName As String = "requestTimesForContractAndTickType"
On Error GoTo Err

Static sRequestTimesForContractAndTickType As Collection

If sRequestTimesForContractAndTickType Is Nothing Then Set sRequestTimesForContractAndTickType = New Collection
Set requestTimesForContractAndTickType = sRequestTimesForContractAndTickType

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function submittedHistDataRequests() As ValueCache
Const ProcName As String = "submittedHistDataRequests"
On Error GoTo Err

Static sSubmittedRequests As ValueCache

If sSubmittedRequests Is Nothing Then Set sSubmittedRequests = CreateValueCache(MaxTwsHistRequestsPerPeriod, "HistRequest")
Set submittedHistDataRequests = sSubmittedRequests

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function




