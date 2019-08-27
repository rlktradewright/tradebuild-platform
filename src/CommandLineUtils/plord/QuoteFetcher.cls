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

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
mDataSource.RemoveGenericTickListener Me
Set mSelfRef = Nothing
End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

If mDataSource.HasCurrentTick(TickTypeAsk) And _
    mDataSource.HasCurrentTick(TickTypeBid) And _
    mDataSource.HasCurrentTick(TickTypeTrade) Then
    Dim lContract As IContract: Set lContract = mDataSource.ContractFuture.Value
    gWriteLineToConsole lContract.Specifier.LocalSymbol & "@" & lContract.Specifier.Exchange & _
                        ": " & _
                        GetCurrentTickSummary(mDataSource)
    mDataSource.RemoveGenericTickListener Me
    Set mSelfRef = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ???? Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub FetchQuote( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractStore As IContractStore, _
                ByVal pMarketDataManager As IMarketDataManager)
Const ProcName As String = "FetchQuote"
On Error GoTo Err

Set mSelfRef = Me

Dim lContractFuture As IFuture
Set lContractFuture = FetchContract(pContractSpec, pContractStore)

Set mDataSource = pMarketDataManager.CreateMarketDataSource(lContractFuture, False)
mDataSource.AddGenericTickListener Me
mDataSource.StartMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




