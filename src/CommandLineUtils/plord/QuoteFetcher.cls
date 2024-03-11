VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QuoteFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IGenericTickListener

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

Private Const ModuleName                            As String = "QuoteFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mSelfRef                                    As Object

Private mDataSource                                 As IMarketDataSource

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContract                                   As IContract

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
writeQuoteMessage
Set mSelfRef = Nothing
End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Select Case ev.Tick.TickType
Case TickTypeBid
Case TickTypeAsk
Case TickTypeTrade
Case TickTypeVolume
Case Else
    Exit Sub
End Select

If gotRequiredTicks Then
    writeQuoteMessage
    mDataSource.Finish
    Set mSelfRef = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    gWriteErrorLine ev.Future.ErrorMessage, ErrorCountIncrementNo
    Set mSelfRef = Nothing
ElseIf ev.Future.IsCancelled Then
    gWriteErrorLine "Contract fetch cancelled", ErrorCountIncrementNo
    Set mSelfRef = Nothing
ElseIf ev.Future.IsAvailable Then
    Set mContract = ev.Future.Value
End If

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

Friend Sub FetchQuote( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractStore As IContractStore, _
                ByVal pMarketDataManager As RealTimeDataManager)
Const ProcName As String = "FetchQuote"
On Error GoTo Err

Set mSelfRef = Me

Set mFutureWaiter = New FutureWaiter
Dim lContractFuture As IFuture
Set lContractFuture = FetchContract(pContractSpec, pContractStore)
mFutureWaiter.Add lContractFuture

Set mDataSource = pMarketDataManager.CreateSnapshotMarketDataSource(lContractFuture, False)

mDataSource.AddGenericTickListener Me
mDataSource.StartMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function gotRequiredTicks() As Boolean
If Not (mDataSource.HasCurrentTick(TickTypeAsk) And _
    mDataSource.HasCurrentTick(TickTypeBid)) Then
    gotRequiredTicks = False
ElseIf mContract.Specifier.SecType <> SecTypeCash And _
        Not (mDataSource.HasCurrentTick(TickTypeTrade) And _
            mDataSource.HasCurrentTick(TickTypeVolume)) Then
    gotRequiredTicks = True
Else
    gotRequiredTicks = False
End If
End Function


Private Sub writeQuoteMessage()
Const ProcName As String = "writeMessage"
On Error GoTo Err

gWriteLineToConsole gPadStringRight(gGetContractName(mContract), 20) & _
                    GetCurrentTickSummary(mDataSource)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
