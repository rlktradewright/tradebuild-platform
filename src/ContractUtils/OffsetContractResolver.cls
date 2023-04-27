VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OffsetContractResolver"
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

Private Const ModuleName                            As String = "OffsetContractResolver"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec                               As IContractSpecifier

Private mExpiryOffset                               As Long
Private mDaysBeforeExpiryToSwitch                   As Long

Private mFetchFuture                                As IFuture

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mFutureBuilder                              As FutureBuilder

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then

ElseIf ev.Future.IsFaulted Then

Else
    Dim lContracts As IContracts: Set lContracts = ev.Future.Value
    
    Dim lContract As IContract
    Dim lIndex As Long: lIndex = -1
    For Each lContract In lContracts
        If gIsContractExpired(lContract) Then
            lIndex = -1
        Else
            lIndex = lIndex + 1
        End If
        If lIndex < mExpiryOffset Then
        ElseIf GetOffsetSessionTimes( _
                            lContract.ExpiryDate, _
                            -mDaysBeforeExpiryToSwitch).StartTime >= Int(CDbl(Now)) _
        Then
            ' we've found the contract we're after
            Exit For
        End If
    Next
            
    If Not lContract Is Nothing Then
        gLogger.Log "Got required offset contract: " & lContract.Specifier.ToString, ProcName, ModuleName, LogLevelDetail
        mFutureBuilder.Value = lContract
        mFutureBuilder.Complete
    Else
        Dim s As String: s = "Can't find contract with offset " & mExpiryOffset & ": " & mContractSpec.ToString
        gLogger.Log s, ProcName, ModuleName, LogLevelDetail
        mFutureBuilder.Fail ErrorCodes.ErrIllegalArgumentException, s, ""
    End If
End If

Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function Resolve( _
                ByVal pSampleContract As IContract, _
                ByVal pExpiryOffset As Long, _
                ByVal pDaysBeforeExpiryToSwitch As Long, _
                ByVal pContractStore As IContractStore) As IFuture
Const ProcName As String = "Resolve"
On Error GoTo Err

Set mSelfRef = Me

mExpiryOffset = pExpiryOffset
mDaysBeforeExpiryToSwitch = pDaysBeforeExpiryToSwitch

' we've got a contract that has all the characteristics of the
' desired contract, except the expiry date may not be correct.
' So now we need to find all the contracts that do so (apart
' from the expiry date)

With pSampleContract.Specifier
    Set mContractSpec = gCreateContractSpecifier("", _
                                                .Symbol, _
                                                .TradingClass, _
                                                .Exchange, _
                                                .SecType, _
                                                .CurrencyCode, _
                                                "", _
                                                .Multiplier, _
                                                .Strike, _
                                                .Right)
End With

gLogger.Log "Fetching all expiries for contract: " & mContractSpec.ToString, ProcName, ModuleName, LogLevelDetail

ReDim lsortkeys(0) As ContractSortKeyIds
lsortkeys(0) = ContractSortKeyExpiry

Set mFetchFuture = gFetchContractsSorted( _
                            mContractSpec, _
                            pContractStore, _
                            lsortkeys, _
                            False)

Set mFutureWaiter = New FutureWaiter
mFutureWaiter.Add mFetchFuture

Set mFutureBuilder = New FutureBuilder
Set Resolve = mFutureBuilder.Future

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




