VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionParametersRequester"
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

Implements IContractDetailsConsumer

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

Private Const ModuleName                            As String = "OptionParametersRequester"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpecifier                          As IContractSpecifier
Private mContractRequester                          As ContractsTwsRequester

Private mContractCache                              As ContractCache

Private mCookie                                     As Variant

Private mRequestType                                As OptionParameterTypes
Private mRequestId                                  As Long

Private mResultFutureBuilder                        As New FutureBuilder
Attribute mResultFutureBuilder.VB_VarHelpID = -1
Private mOptionParametersFutureBuilder              As FutureBuilder
Attribute mOptionParametersFutureBuilder.VB_VarHelpID = -1

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mExchange                                   As String
Private mExpiry                                     As String
Private mStrike                                     As Double

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
GIBEnhApi.Logger.Log "OptionParametersRequester terminated for: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
End Sub

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContracts(ByVal pRequestId As Long)

End Sub

Private Sub IContractDetailsConsumer_EndOptionParameters(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndOptionParameters"
On Error GoTo Err

mOptionParametersFutureBuilder.Value = Nothing
mOptionParametersFutureBuilder.Complete

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, ByVal pTwsContract As TwsContract)

End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

mContractRequester.CancelRequest mRequestId

mOptionParametersFutureBuilder.Fail pErrorCode, pErrorMsg, ""
mResultFutureBuilder.Fail pErrorCode, pErrorMsg, ""

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyMarketRule(ByVal pMarketRuleId As Long, pPriceIncrement() As TwsPriceIncrement)

End Sub

Private Sub IContractDetailsConsumer_NotifyOptionParameter( _
                ByVal pRequestId As Long, _
                ByVal pExchange As String, _
                ByVal pUnderlyingContractId As Long, _
                ByVal pTradingClass As String, _
                ByVal pMultiplier As Double, _
                pExpirations() As String, _
                pStrikes() As Double)
Const ProcName As String = "IContractDetailsConsumer_NotifyOptionParameter"
On Error GoTo Err

mContractCache.AddOptionParameter mContractSpecifier, pExchange, pExpirations, pStrikes

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

Select Case mRequestType
Case OptionParameterTypeExpiries
    Dim lExpiries As Expiries
    If GetExpiries(lExpiries) Then
        GIBEnhApi.Logger.Log "Option Expiries retrieved from cache for Contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        mResultFutureBuilder.Value = lExpiries
    Else
        GIBEnhApi.Logger.Log "No Option Expiries retrieved from cache for Contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
    End If
Case OptionParameterTypeStrikes
    Dim lStrikes As Strikes
    If GetStrikes(lStrikes) Then
        GIBEnhApi.Logger.Log "Option Strikes retrieved from cache for Contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        mResultFutureBuilder.Value = lStrikes
    Else
        GIBEnhApi.Logger.Log "No Option Strikes retrieved from cache for Contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
    End If
End Select

mResultFutureBuilder.Complete

Set mSelfRef = Nothing

Exit Sub

Err:
GIBEnhApi.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function Fetch( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pRequestType As OptionParameterTypes, _
                ByVal pExchange As String, _
                ByVal pExpiry As String, _
                ByVal pStrike As Double, _
                ByVal pCookie As Variant) As IFuture
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mContractSpecifier = pContractSpecifier
mRequestType = pRequestType
mExchange = pExchange
mExpiry = pExpiry
mStrike = pStrike
GIBEnhApi.SetVariant mCookie, pCookie

Select Case mRequestType
Case OptionParameterTypeExpiries
    Dim lExpiries As Expiries
    If GetExpiries(lExpiries) Then
        GIBEnhApi.Logger.Log "Option Expiries retrieved from cache for Contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        Set Fetch = CreateFuture(lExpiries)
        Exit Function
    End If
    
    Set mSelfRef = Me
    
    If mContractCache.GetOptionParametersFutureBuilder(mContractSpecifier, mOptionParametersFutureBuilder) Then
        GIBEnhApi.Logger.Log "Awaiting Option Expiries from cache for Contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        mFutureWaiter.Add mOptionParametersFutureBuilder.Future
        Set Fetch = mResultFutureBuilder.Future
        Exit Function
    End If
    
    GIBEnhApi.Logger.Log "Fetching Option Expiries from TWS for contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
    Set mOptionParametersFutureBuilder = New FutureBuilder
    mContractCache.SetOptionParametersFutureBuilder mContractSpecifier, mOptionParametersFutureBuilder
    mFutureWaiter.Add mOptionParametersFutureBuilder.Future
    Set Fetch = mResultFutureBuilder.Future
Case OptionParameterTypeStrikes
    Dim lStrikes As Strikes
    If GetStrikes(lStrikes) Then
        GIBEnhApi.Logger.Log "Option Strikes retrieved from cache for contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        Set Fetch = CreateFuture(lStrikes)
        Exit Function
    End If
    
    Set mSelfRef = Me
    
    If mContractCache.GetOptionParametersFutureBuilder(mContractSpecifier, mOptionParametersFutureBuilder) Then
        GIBEnhApi.Logger.Log "Awaiting Option Strikes from cache for contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        mFutureWaiter.Add mOptionParametersFutureBuilder.Future
        Set Fetch = mResultFutureBuilder.Future
        Exit Function
    End If
    
    GIBEnhApi.Logger.Log "Fetching Option Strikes from TWS for contract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
    Set mOptionParametersFutureBuilder = New FutureBuilder
    mContractCache.SetOptionParametersFutureBuilder mContractSpecifier, mOptionParametersFutureBuilder
    mFutureWaiter.Add mOptionParametersFutureBuilder.Future
    Set Fetch = mResultFutureBuilder.Future
End Select

mRequestId = mContractRequester.RequestOptionParameters(mContractSpecifier, Me)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================
 
Private Function GetExpiries(ByRef pExpiries As Expiries) As Boolean
Const ProcName As String = "getExpiries"
On Error GoTo Err

GetExpiries = mContractCache.GetOptionExpiries(mContractSpecifier, mExchange, mStrike, pExpiries)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetStrikes(ByRef pStrikes As Strikes) As Boolean
Const ProcName As String = "getStrikes"
On Error GoTo Err

GetStrikes = mContractCache.GetOptionStrikes(mContractSpecifier, mExchange, mExpiry, pStrikes)

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function


