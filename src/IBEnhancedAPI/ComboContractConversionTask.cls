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
Implements Task

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

Private mContractSpec                               As ContractSpecifier

Private mContractRequester                          As ContractDetailsRequester

Private mRequestIds()                               As Long

Private mErrorCode                                  As Long
Private mErrorMsg                                   As String

Private mRequestsOutstanding                        As Long

Private mLegs                                       As New EnumerableCollection

Private mTimezoneId                                 As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IContractDetailsConsumer Interface Members
'@================================================================================

Private Sub IContractDetailsConsumer_EndContractDetails(ByVal pRequestId As Long)
Const ProcName As String = "IContractDetailsConsumer_EndContractDetails"
On Error GoTo Err

mRequestsOutstanding = mRequestsOutstanding - 1
If mRequestsOutstanding = 0 Then
    If mErrorCode <> 0 Then
        mTaskContext.Error mErrorCode, mErrorMsg
    Else
        mTaskContext.Finish generateTwsContractDetails, False
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IContractDetailsConsumer_NotifyContract(ByVal pRequestId As Long, pContractDetails As TwsContractDetails)
Const ProcName As String = "IContractDetailsConsumer_NotifyContract"
On Error GoTo Err

If mErrorCode <> 0 Then Exit Sub
    
If alreadyGotContract(pRequestId) Then
    mErrorCode = 200
    mErrorMsg = "Contract not uniquely specified for leg " & CStr(getLegIndexFromRequestId(pRequestId) + 1)
Else
    mLegs.Add createTwsComboLeg(pContractDetails, mContractSpec.ComboLegs(getLegIndexFromRequestId(pRequestId) + 1)), CStr(pRequestId)
    If mTimezoneId = "" And pContractDetails.TimeZoneId <> "" Then mTimezoneId = pContractDetails.TimeZoneId
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_Run()
Const ProcName As String = "Task_Run"
On Error GoTo Err

Dim lComboLeg As comboLeg

For Each lComboLeg In mContractSpec.ComboLegs
    mRequestIds(mRequestsOutstanding) = mContractRequester.RequestContractDetails(lComboLeg.ContractSpec, Me)
    mRequestsOutstanding = mRequestsOutstanding + 1
Next

mTaskContext.Suspend -1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Let Task_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
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
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractSpec As ContractSpecifier)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pContractSpec.Sectype <> SecTypeCombo Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Argument is not a combo contract specifier"
If pContractSpec.ComboLegs.Count = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "No combo legs specified"

Set mContractRequester = pContractRequester
Set mContractSpec = pContractSpec

ReDim mRequestIds(mContractSpec.ComboLegs.Count - 1) As Long

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function createTwsComboLeg( _
                ByVal pContract As TwsContract, _
                ByVal pComboLeg As comboLeg) As TwsComboLeg
Const ProcName As String = "createTwsComboLeg"
On Error GoTo Err

Set createTwsComboLeg = New TwsComboLeg

createTwsComboLeg.ConId = pContract.ConId
createTwsComboLeg.Action = IIf(pComboLeg.IsBuyLeg, TwsOrderActionBuy, TwsOrderActionSell)
createTwsComboLeg.Exchange = pContract.Exchange
createTwsComboLeg.Ratio = pComboLeg.Ratio
createTwsComboLeg.OpenClose = TwsLegOpenCloseSame
createTwsComboLeg.ShortSaleSlot = TwsShortSaleSlotNotApplicable

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function generateTwsContractDetails() As TwsContractDetails
Const ProcName As String = "generateTwsContractDetails"
On Error GoTo Err

Dim i As Long
Dim lContract As TwsContract

Set lContract = New TwsContract

lContract.Symbol = mContractSpec.Symbol
lContract.Sectype = TwsSecTypeCombo
lContract.Exchange = mContractSpec.Exchange
lContract.CurrencyCode = mContractSpec.CurrencyCode

For i = 0 To UBound(mRequestIds)
    lContract.ComboLegs.Add mLegs(CStr(mRequestIds(i)))
Next

Set generateTwsContractDetails = New TwsContractDetails
generateTwsContractDetails.Summary = lContract
generateTwsContractDetails.TimeZoneId = mTimezoneId

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
Err.Raise ErrorCodes.ErrIllegalStateException, , "Invalid request id"

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function


