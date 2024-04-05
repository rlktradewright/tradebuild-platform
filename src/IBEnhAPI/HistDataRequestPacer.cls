VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataRequestPacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "HistDataRequestPacer"

Private Const OneMinute                             As Double = 1 / (60# * 24#)
Private Const OneSecond                             As Double = 1 / (60# * 60# * 24#)

Private Const MaxTwsHistRequestsPerPeriod           As Long = 60

Private Const TwsHistRequestPeriod                  As Double = 10# * OneMinute + 2# * OneSecond

Private Const TwsHistRequestRepeatInterval          As Double = 15# * OneSecond

Private Const TwsHistRequestContractAndTickTypePeriod       As Double = 10# * OneSecond
Private Const TwsHistRequestContractAndTickTypeRequestLimit As Long = 6

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

Public Function GetDelayTillNextSubmission(ByRef pReason As String) As Long
Const ProcName As String = "GetDelayTillNextSubmission"
On Error GoTo Err

If submittedHistDataRequests.Count < submittedHistDataRequests.CyclicSize Then Exit Function
    
Dim lTimestamp As Date
lTimestamp = GetTimestamp
    
Dim lEarliestSubmissionTime As Date
lEarliestSubmissionTime = submittedHistDataRequests.GetSValue(1).Timestamp + TwsHistRequestPeriod
If lTimestamp < lEarliestSubmissionTime Then
    GetDelayTillNextSubmission = ((lEarliestSubmissionTime - lTimestamp) * 86400) * 1000
    pReason = "overall pacing"
Else
    GetDelayTillNextSubmission = 0
End If

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetEarliestSubmissionTime( _
                ByRef pRequest As TwsHistoricalDataRequest, _
                ByRef pReason As String) As Date
Const ProcName As String = "GetEarliestSubmissionTime"
On Error GoTo Err

pReason = ""

GetEarliestSubmissionTime = getEarliestTimeForIdenticalRequest(pRequest)
If GetEarliestSubmissionTime <> 0 Then pReason = "identical request: " & historicalDataRequestToString(pRequest)

Dim lContractTime As Date
lContractTime = getEarliestTimeForContractAndTickType(pRequest)

If lContractTime > GetEarliestSubmissionTime Then
    GetEarliestSubmissionTime = lContractTime
    pReason = "same contract and tick type: " & historicalDataRequestToString(pRequest)
End If

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function RecordSubmission(ByRef pRequest As TwsHistoricalDataRequest) As Date
Const ProcName As String = "RecordSubmission"
On Error GoTo Err


Dim lTimestamp As Date
lTimestamp = GetTimestamp

submittedHistDataRequests.AddValue pRequest, 0, lTimestamp, 0

Dim lKey As String
lKey = historicalDataRequestToString(pRequest)
If requestTimes.Contains(lKey) Then requestTimes.Remove lKey

requestTimes.Add lTimestamp, lKey

Dim lCache As ValueCache

Set lCache = getContractAndTickTypeCache(pRequest)
lCache.AddValue pRequest, 0, lTimestamp, 0

RecordSubmission = lTimestamp

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function contractAndTickTypeToString(ByRef pRequest As TwsHistoricalDataRequest) As String
Const ProcName As String = "contractAndTickTypeToString"
On Error GoTo Err

contractAndTickTypeToString = pRequest.ContractSpec.ToString & vbNullChar & _
                                pRequest.WhatToShow

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getContractAndTickTypeCache(ByRef pRequest As TwsHistoricalDataRequest) As ValueCache
Const ProcName As String = "getContractAndTickTypeCache"
On Error GoTo Err

Dim lKey As String
lKey = contractAndTickTypeToString(pRequest)

Dim lCache As ValueCache
If Not requestTimesForContractAndTickType.TryItem(lKey, lCache) Then
    Set lCache = CreateValueCache(TwsHistRequestContractAndTickTypeRequestLimit - 1, "Time")
    requestTimesForContractAndTickType.Add lCache, lKey
End If

Set getContractAndTickTypeCache = lCache

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getEarliestTimeForContractAndTickType(ByRef pRequest As TwsHistoricalDataRequest) As Date
Const ProcName As String = "getEarliestTimeForContractAndTickType"
On Error GoTo Err

Dim lCache As ValueCache
Set lCache = getContractAndTickTypeCache(pRequest)

If lCache.Count = lCache.CyclicSize Then
    getEarliestTimeForContractAndTickType = lCache.GetSValue(1).Timestamp + TwsHistRequestContractAndTickTypePeriod
End If

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getEarliestTimeForIdenticalRequest(ByRef pRequest As TwsHistoricalDataRequest) As Date
Const ProcName As String = "getEarliestTimeForIdenticalRequest"
On Error GoTo Err

Dim lLastTime As Date

Dim lKey As String
lKey = historicalDataRequestToString(pRequest)
If requestTimes.Contains(lKey) Then
    lLastTime = CDate(requestTimes(lKey))
    getEarliestTimeForIdenticalRequest = lLastTime + TwsHistRequestRepeatInterval
End If

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function historicalDataRequestToString(ByRef pRequest As TwsHistoricalDataRequest) As String
Const ProcName As String = "historicalDataRequestToString"
On Error GoTo Err

historicalDataRequestToString = pRequest.BarSizeSetting & vbNullChar & _
                                pRequest.ContractSpec.ToString & vbNullChar & _
                                pRequest.Duration & vbNullChar & _
                                pRequest.EndDateTime & vbNullChar & _
                                pRequest.WhatToShow

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function requestTimes() As EnumerableCollection
Const ProcName As String = "requestTimes"
On Error GoTo Err

Static sRequestTimes As EnumerableCollection

If sRequestTimes Is Nothing Then Set sRequestTimes = New EnumerableCollection
Set requestTimes = sRequestTimes

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function requestTimesForContractAndTickType() As EnumerableCollection
Const ProcName As String = "requestTimesForContractAndTickType"
On Error GoTo Err

Static sRequestTimesForContractAndTickType As EnumerableCollection

If sRequestTimesForContractAndTickType Is Nothing Then Set sRequestTimesForContractAndTickType = New EnumerableCollection
Set requestTimesForContractAndTickType = sRequestTimesForContractAndTickType

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function submittedHistDataRequests() As ValueCache
Const ProcName As String = "submittedHistDataRequests"
On Error GoTo Err

Static sSubmittedRequests As ValueCache

If sSubmittedRequests Is Nothing Then Set sSubmittedRequests = CreateValueCache(MaxTwsHistRequestsPerPeriod - 1, "HistRequest")
Set submittedHistDataRequests = sSubmittedRequests

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function






