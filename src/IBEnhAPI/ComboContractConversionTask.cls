VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboContractConversionTask"
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

Private Const ModuleName                            As String = "ComboContractConversionTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mContractSpec                               As IContractSpecifier

Private mContractRequester                          As ContractsTwsRequester

Private mRequestIds()                               As Long

Private mErrorCode                                  As Long
Private mErrorMsg                                   As String

Private mRequestsOutstanding                        As Long

Private mLegs                                       As New EnumerableCollection

Private mTimeZoneId                                 As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContracts(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContracts"
On Error GoTo Err

mRequestsOutstanding = mRequestsOutstanding - 1
If mRequestsOutstanding = 0 Then
    If mErrorCode <> 0 Then
        mTaskContext.Error mErrorCode, mErrorMsg
    Else
        mTaskContext.Finish generateTwsContract, False
    End If
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_EndOptionParameters(ByVal pRequestIs As Long)

End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, ByVal pTwsContract As TwsContract)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

If mErrorCode <> 0 Then Exit Sub
    
If alreadyGotContract(pRequestId) Then
    mErrorCode = 200
    mErrorMsg = "Contract not uniquely specified for leg " & CStr(getLegIndexFromRequestId(pRequestId) + 1)
Else
    mLegs.Add createTwsComboLeg(pTwsContract.Specifier, mContractSpec.ComboLegs(getLegIndexFromRequestId(pRequestId) + 1)), CStr(pRequestId)
    If mTimeZoneId = "" And pTwsContract.TimeZoneId <> "" Then mTimeZoneId = pTwsContract.TimeZoneId
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyError(ByVal pRequestId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IContractDetailsConsumer_NotifyError"
On Error GoTo Err

mRequestsOutstanding = mRequestsOutstanding - 1

If mErrorCode = 0 Then
    mErrorCode = pErrorCode
    mErrorMsg = pErrorMsg
End If

If mRequestsOutstanding = 0 Then mTaskContext.Error mErrorCode, mErrorMsg

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyMarketRule(ByVal pMarketRuleId As Long, pPriceIncrement() As TwsPriceIncrement)

End Sub

Private Sub IContractDetailsConsumer_NotifyOptionParameter(ByVal pRequestId As Long, ByVal pExchange As String, ByVal pUnderlyingContractId As Long, ByVal pTradingClass As String, ByVal pMultiplier As Double, pExpirations() As String, pStrikes() As Double)

End Sub

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Dim lComboLeg As comboLeg

For Each lComboLeg In mContractSpec.ComboLegs
    mRequestIds(mRequestsOutstanding) = mContractRequester.RequestContractDetails(lComboLeg.ContractSpec, Me)
    mRequestsOutstanding = mRequestsOutstanding + 1
Next

mTaskContext.Suspend -1

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument pContractSpec.Sectype = SecTypeCombo, "Argument is not a combo contract specifier"
AssertArgument pContractSpec.ComboLegs.Count <> 0, "No combo legs specified"

Set mContractRequester = pContractRequester
Set mContractSpec = pContractSpec

ReDim mRequestIds(mContractSpec.ComboLegs.Count - 1) As Long

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function alreadyGotContract(ByVal pRequestId As Long) As Boolean
Const ProcName As String = "alreadyGotContract"
On Error GoTo Err

alreadyGotContract = mLegs.Contains(CStr(pRequestId))

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function createTwsComboLeg( _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pComboLeg As comboLeg) As TwsComboLeg
Const ProcName As String = "createTwsComboLeg"
On Error GoTo Err

Set createTwsComboLeg = New TwsComboLeg

createTwsComboLeg.ConId = pContractSpec.ConId
createTwsComboLeg.Action = IIf(pComboLeg.IsBuyLeg, TwsOrderActionBuy, TwsOrderActionSell)
createTwsComboLeg.Exchange = pContractSpec.Exchange
createTwsComboLeg.Ratio = pComboLeg.Ratio
createTwsComboLeg.OpenClose = TwsLegOpenCloseSame
createTwsComboLeg.ShortSaleSlot = TwsShortSaleSlotNotApplicable

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateTwsContract() As TwsContract
Const ProcName As String = "generateTwsContract"
On Error GoTo Err

Dim i As Long
Dim lContractSpec As TwsContractSpecifier

Set lContractSpec = New TwsContract

lContractSpec.Symbol = mContractSpec.Symbol
lContractSpec.Sectype = TwsSecTypeCombo
lContractSpec.Exchange = mContractSpec.Exchange
lContractSpec.CurrencyCode = mContractSpec.CurrencyCode

For i = 0 To UBound(mRequestIds)
    lContractSpec.ComboLegs.Add mLegs(CStr(mRequestIds(i)))
Next

Set generateTwsContract = New TwsContract
generateTwsContract.Specifier = lContractSpec
generateTwsContract.TimeZoneId = mTimeZoneId

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getLegIndexFromRequestId(ByVal pRequestId As Long) As Long
Dim i As Long
Const ProcName As String = "getLegIndexFromRequestId"
On Error GoTo Err

For i = 0 To UBound(mRequestIds)
    If mRequestIds(i) = pRequestId Then
        getLegIndexFromRequestId = i
        Exit Function
    End If
Next
Assert False, "Invalid request id"

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function


