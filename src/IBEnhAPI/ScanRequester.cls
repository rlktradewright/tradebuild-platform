VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScanRequester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IScannerDataConsumer

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

Private Const ModuleName                            As String = "ScanRequester"

'@================================================================================
' Member variables
'@================================================================================

Private mTwsAPI                                     As TwsAPI

Private mNextRequestID                              As Long

Private mRequests                                   As EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mRequests = New EnumerableCollection
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

Private Sub IScannerDataConsumer_EndScannerData(ByVal pRequestId As Long)
Const ProcName As String = "IScannerDataConsumer_EndScannerData"
On Error GoTo Err

Dim lRequest As ScanRequest
lRequest = mRequests.Item(CStr(pRequestId))

lRequest.FutureBuilder.Value = lRequest.Results
lRequest.FutureBuilder.Complete

If lRequest.AutoRefresh Then
    Set lRequest.FutureBuilder = New FutureBuilder
Else
    mTwsAPI.CancelScannerSubscription lRequest.requestId
    mRequests.Remove CStr(pRequestId)
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IScannerDataConsumer_NotifyError( _
                ByVal pRequestId As Long, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "IScannerDataConsumer_NotifyError"
On Error GoTo Err

Dim lRequest As ScanRequest
' ignore "API scanner subscription cancelled" messages which arrive
' after the requester has finished
If Not mRequests.TryItem(CStr(pRequestId), lRequest) Then Exit Sub

If pErrorCode = 165 Then
    ' 165 = no items retrieved
    ' ignore this as it is followed by EndScannerData
    Exit Sub
End If

lRequest.FutureBuilder.Fail pErrorCode, pErrorMsg, ProcName
mRequests.Remove CStr(pRequestId)


Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IScannerDataConsumer_NotifyScannerData( _
                ByVal pRequestId As Long, _
                ByVal pRank As Long, _
                ByVal pContract As TwsContract, _
                ByVal pDistance As String, _
                ByVal pBenchmark As String, _
                ByVal pProjection As String, _
                ByVal pLegs As String)
Const ProcName As String = "IScannerDataConsumer_NotifyScannerData"
On Error GoTo Err

Dim lResult As New ScanResult
lResult.Contract = GIBEnhancedApi.TwsContractToContract(pContract)
lResult.Rank = pRank
lResult.AddAttribute "Benchmark", pBenchmark
lResult.AddAttribute "Distance", pDistance
lResult.AddAttribute "Legs", pLegs
lResult.AddAttribute "Projection", pProjection

Dim lRequest As ScanRequest
lRequest = mRequests.Item(CStr(pRequestId))
lRequest.Results.Add lResult

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IScannerDataConsumer_NotifyScannerParameters(ByVal pXml As String)
Assert False, "Not implemented", ErrorCodes.ErrUnsupportedOperationException
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub CancelScan( _
                ByVal pCookie As String)
Const ProcName As String = "CancelScan"
On Error GoTo Err

Dim en As Enumerator: Set en = mRequests.Enumerator
Do While en.MoveNext
    Dim lRequest As ScanRequest: lRequest = en.Current
    If lRequest.Cookie = pCookie Then
        mTwsAPI.CancelScannerSubscription lRequest.requestId
        en.Remove
        Exit Do
    End If
Loop

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI)
Set mTwsAPI = pTwsAPI
End Sub

Friend Function RefreshData( _
                ByVal pCookie As String) As IFuture
Const ProcName As String = "RefreshData"
On Error GoTo Err

Dim en As Enumerator: Set en = mRequests.Enumerator
Do While en.MoveNext
    Dim lRequest As ScanRequest: lRequest = en.Current
    If lRequest.Cookie = pCookie Then
        Set RefreshData = lRequest.FutureBuilder.Future
        Exit Do
    End If
Loop

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function ScanMarket( _
                ByVal pScanName As String, _
                ByVal pParameters As Parameters, _
                ByVal pOptions As Parameters, _
                ByVal pFilters As Parameters, _
                ByVal pAutoRefresh As Boolean, _
                ByVal pCookie As String) As IFuture
Const ProcName As String = "ScanMarket"
On Error GoTo Err

Dim lSubscription As New TwsScannerSubscription
With lSubscription
    .ScanCode = UCase$(pScanName)
    
    .NumberOfRows = pParameters.GetParameterValue("NumberOfRows", CStr(&H7FFFFFFF))
    .Instrument = UCase$(pParameters.GetParameterValue("Instrument", ""))
    .LocationCode = UCase$(pParameters.GetParameterValue("LocationCode", ""))
    .AbovePrice = updateDoubleFromParameters(pParameters, "AbovePrice", MaxDouble)
    .BelowPrice = updateDoubleFromParameters(pParameters, "BelowPrice", MaxDouble)
    .AboveVolume = pParameters.GetParameterValue("AboveVolume", CStr(&H7FFFFFFF))
    .AverageOptionVolumeAbove = pParameters.GetParameterValue("AverageOptionVolumeAbove", CStr(&H7FFFFFFF))
    .MarketCapAbove = updateDoubleFromParameters(pParameters, "MarketCapAbove", MaxDouble)
    .MarketCapBelow = updateDoubleFromParameters(pParameters, "MarketCapBelow", MaxDouble)
    .MoodyRatingAbove = UCase$(pParameters.GetParameterValue("MoodyRatingAbove", ""))
    .MoodyRatingBelow = UCase$(pParameters.GetParameterValue("MoodyRatingBelow", ""))
    .SpRatingAbove = UCase$(pParameters.GetParameterValue("SpRatingAbove", ""))
    .SpRatingBelow = UCase$(pParameters.GetParameterValue("SpRatingBelow", ""))
    .MaturityDateAbove = pParameters.GetParameterValue("MaturityDateAbove", "")
    .MaturityDateBelow = pParameters.GetParameterValue("MaturityDateBelow", "")
    .CouponRateAbove = updateDoubleFromParameters(pParameters, "CouponRateAbove", MaxDouble)
    .CouponRateBelow = updateDoubleFromParameters(pParameters, "CouponRateBelow", MaxDouble)
    .ExcludeConvertible = pParameters.GetParameterValue("ExcludeConvertible", "False")
    .ScannerSettingPairs = UCase$(pParameters.GetParameterValue("ScannerSettingPairs", ""))
    .StockTypeFilter = UCase$(pParameters.GetParameterValue("StockTypeFilter", ""))
End With

Dim lScanRequest As ScanRequest
lScanRequest.AutoRefresh = pAutoRefresh
lScanRequest.Cookie = pCookie
Set lScanRequest.FutureBuilder = New FutureBuilder
lScanRequest.requestId = mNextRequestID
mNextRequestID = mNextRequestID + 1
Set lScanRequest.Results = New ScanResults
mRequests.Add lScanRequest, CStr(lScanRequest.requestId)

mTwsAPI.RequestScannerSubscription lScanRequest.requestId, _
                                    lSubscription, _
                                    pOptions, _
                                    pFilters

Set ScanMarket = lScanRequest.FutureBuilder.Future

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrTypeMismatch Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "Parameter type mismatch"
Else
    GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function updateDoubleFromParameters( _
                ByVal pParams As Parameters, _
                ByVal pParamName As String, _
                ByVal pDefault As Double) As Double
Const ProcName As String = "updateDoubleFromParameters"
On Error GoTo Err

If pParams.GetParameterValue(pParamName, "") <> "" Then
    updateDoubleFromParameters = CDbl(pParams.GetParameterValue(pParamName))
Else
    updateDoubleFromParameters = pDefault
End If

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function


