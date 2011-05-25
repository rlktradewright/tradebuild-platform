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

Public Property Get DelayTillNextSubmission() As Long
Const ProcName As String = "DelayTillNextSubmission"
On Error GoTo Err

Dim lTimestamp As Date
Dim lEarliestSubmissionTime As Date

lTimestamp = GetTimestamp

If submittedHistDataRequests.Count = submittedHistDataRequests.CyclicSize Then
    lEarliestSubmissionTime = submittedHistDataRequests.GetSValue(1).Timestamp + TwsHistRequestPeriod
    If lTimestamp < lEarliestSubmissionTime Then
        DelayTillNextSubmission = ((lEarliestSubmissionTime - lTimestamp) * 86400 + 5) * 1000
    Else
        DelayTillNextSubmission = 0
    End If
Else
    DelayTillNextSubmission = 0
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function GetEarliestSubmissionTime(ByRef pRequest As HistoricalDataRequest) As Date
Const ProcName As String = "GetEarliestSubmissionTime"
On Error GoTo Err

GetEarliestSubmissionTime = getEarliestTimeForIdenticalRequest(pRequest)

Dim lContractTime As Date
lContractTime = getEarliestTimeForContractAndTickType(pRequest)

If lContractTime > GetEarliestSubmissionTime Then GetEarliestSubmissionTime = lContractTime

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function RecordSubmission(ByRef pRequest As HistoricalDataRequest) As Date
Const ProcName As String = "RecordSubmission"
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

RecordSubmission = lTimestamp

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

contractAndTickTypeToString = pRequest.Contract.ToString & vbNullChar & _
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

historicalDataRequestToString = pRequest.BarSizeSetting & vbNullChar & _
                                pRequest.Contract.ToString & vbNullChar & _
                                pRequest.Duration & vbNullChar & _
                                pRequest.EndDateTime & vbNullChar & _
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






