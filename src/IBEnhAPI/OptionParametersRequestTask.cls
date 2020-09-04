VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionParametersRequestTask"
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
Implements ITask

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

Private Const ModuleName                            As String = "OptionParametersRequestTask"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpecifier                          As IContractSpecifier
Private mContractRequester                          As ContractsTwsRequester

Private mTaskContext                                As TaskContext

Private mContractCache                              As ContractCache

Private mCookie                                     As Variant

Private mRequestType                                As OptionParameterTypes
Private mRequestId                                  As Long

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1

Private mExchange                                   As String
Private mExpiry                                     As String
Private mStrike                                     As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureBuilder = New FutureBuilder
End Sub

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContracts(ByVal pRequestId As Long)

End Sub

Private Sub IContractDetailsConsumer_EndOptionParameters(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndOptionParameters"
On Error GoTo Err

mContractCache.EndOptionParameters mContractSpecifier

Select Case mRequestType
Case OptionParameterTypeExpiries
    If GetExpiries Then
        gLogger.Log "Option Expiries retrieved from Tws for IContract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        Exit Sub
    End If
Case OptionParameterTypeStrikes
    If GetStrikes Then
        gLogger.Log "Option Strikes retrieved from Tws for IContract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        Exit Sub
    End If
End Select

mTaskContext.Finish Nothing, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, ByVal pTwsContract As TwsContract)

End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

mContractRequester.CancelRequest mRequestId
mTaskContext.Error pErrorCode, pErrorMsg

mFutureBuilder.Fail pErrorCode, pErrorMsg, ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyMarketRule(ByVal pMarketRuleId As Long, pPriceIncrement() As IBAPIV100.TwsPriceIncrement)

End Sub

Private Sub IContractDetailsConsumer_NotifyOptionParameter(ByVal pRequestId As Long, ByVal pExchange As String, ByVal pUnderlyingContractId As Long, ByVal pTradingClass As String, ByVal pMultiplier As Long, pExpirations() As String, pStrikes() As Double)
Const ProcName As String = "IContractDetailsConsumer_NotifyOptionParameter"
On Error GoTo Err

mContractCache.AddOptionParameter mContractSpecifier, pExchange, pExpirations, pStrikes

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
doCancel
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

If mTaskContext.CancelPending Then
    doCancel
    Exit Sub
End If

Select Case mRequestType
Case OptionParameterTypeExpiries
    If GetExpiries Then
        gLogger.Log "Option Expiries retrieved from cache for IContract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        Exit Sub
    End If
Case OptionParameterTypeStrikes
    If GetStrikes Then
        gLogger.Log "Option Strikes retrieved from cache for IContract: " & mContractSpecifier.ToString, ProcName, ModuleName, LogLevelDetail
        Exit Sub
    End If
End Select

mRequestId = mContractRequester.RequestOptionParameters(mContractSpecifier, Me)

mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mFutureBuilder Event Handlers
'@================================================================================

Private Sub mFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mFutureBuilder_Cancelled"
On Error GoTo Err

doCancel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Future() As IFuture
Set Future = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractCache As ContractCache, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pRequestType As OptionParameterTypes, _
                ByVal pExchange As String, _
                ByVal pExpiry As String, _
                ByVal pStrike As Double, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
Set mContractSpecifier = pContractSpecifier
mRequestType = pRequestType
mExchange = pExchange
mExpiry = pExpiry
mStrike = pStrike
gSetVariant mCookie, pCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doCancel()
Const ProcName As String = "doCancel"
On Error GoTo Err

mContractRequester.CancelRequest mRequestId

mTaskContext.Finish Nothing, True

mFutureBuilder.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function GetExpiries() As Boolean
Const ProcName As String = "getExpiries"
On Error GoTo Err

Dim lExpiries As Expiries
Set lExpiries = mContractCache.GetOptionExpiries(mContractSpecifier, mExchange, mStrike)
If Not lExpiries Is Nothing Then
    mFutureBuilder.Value = lExpiries
    mFutureBuilder.Complete
    mTaskContext.Finish Empty, False
    GetExpiries = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetStrikes() As Boolean
Const ProcName As String = "getStrikes"
On Error GoTo Err

Dim lStrikes As Strikes
Set lStrikes = mContractCache.GetOptionStrikes(mContractSpecifier, mExchange, mExpiry)
If Not lStrikes Is Nothing Then
    mFutureBuilder.Value = lStrikes
    mFutureBuilder.Complete
    mTaskContext.Finish Empty, False
    GetStrikes = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

