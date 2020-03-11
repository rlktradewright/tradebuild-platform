VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractResolver"
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

Implements IStateChangeListener

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

Private Const ModuleName                            As String = "ContractResolver"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private WithEvents mContractProcessor               As ContractProcessor
Attribute mContractProcessor.VB_VarHelpID = -1

Private mBatchOrders                                As Boolean
Private mStageOrders                                As Boolean

Private mContractStore                              As IContractStore

Private mContractSpec                               As IContractSpecifier

Private mMaxExpenditure                             As Long
Private mUnderlyingExchangeName                     As String

Private mFutureBuilder                              As New FutureBuilder

Private mPendingCommands                            As SortedDictionary

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingCommands = CreateSortedDictionary(KeyTypeDate)
End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Static sInputPauseEnded As Boolean

Dim lState As OptionContractResolverStates: lState = ev.State
Select Case lState
Case OptionContractResolverStateWaitingForUnderlyingMarketPrice, _
        OptionContractResolverStateFetchingInitialContract
    ' the underlying contract has been found (or was not needed)
    If Not sInputPauseEnded Then
        sInputPauseEnded = True
        gInputPaused = False
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mContractProcessor Event Handlers
'@================================================================================

Private Sub mContractProcessor_IsReady()
Const ProcName As String = "mContractProcessor_IsReady"
On Error GoTo Err

processPendingCommands
Set mContractProcessor = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    gWriteErrorLine "Contract fetch was cancelled", True
    gInputPaused = False
    Set mSelfRef = Nothing
    Exit Sub
End If

If ev.Future.IsFaulted Then
    gWriteErrorLine ev.Future.ErrorMessage, True
    gInputPaused = False
    Set mSelfRef = Nothing
    Exit Sub
End If

processTargetContract ev.Future.Value
gInputPaused = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ContractFuture() As IFuture
Set ContractFuture = mFutureBuilder.Future
End Property

Friend Property Get ContractName() As String
With mContractSpec
    Dim lStrikeStr As String
    If mMaxExpenditure = 0 Then
        If .Strike <> 0 Then lStrikeStr = CStr(.Strike)
    Else
        lStrikeStr = mMaxExpenditure & "$" & _
                    IIf(mUnderlyingExchangeName <> "", "," & mUnderlyingExchangeName, "")
    End If
    
    ContractName = IIf(.Symbol <> "", .Symbol, "") & _
                    IIf(.LocalSymbol <> "", ";" & .LocalSymbol, "") & _
                    ";" & SecTypeToShortString(.SecType) & _
                    IIf(.Exchange <> "", ";" & .Exchange, "") & _
                    IIf(.CurrencyCode <> "", ";" & .CurrencyCode, "") & _
                    IIf(.Expiry <> "", ";" & .Expiry, "") & _
                    IIf(lStrikeStr <> "", ";" & lStrikeStr, "") & _
                    IIf(.Right <> OptNone, ";" & OptionRightToString(.Right), "") & _
                    IIf(.Multiplier <> 1, ";" & .Multiplier, "")
End With

End Property

Friend Property Get MaxExpenditure() As Long
MaxExpenditure = mMaxExpenditure
End Property

Friend Property Get UnderlyingExchangeName() As String
UnderlyingExchangeName = mUnderlyingExchangeName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddPendingCommand( _
                ByVal pCommand As Command, _
                ByVal pParams As String, _
                Optional ByVal pMessage As String)
Dim lCommand As New PendingCommand
Set lCommand.Command = pCommand
lCommand.Params = pParams
mPendingCommands.Add lCommand, GetTimestamp
If pMessage <> "" Then gWriteLineToConsole pMessage, True
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mContractProcessor = Nothing
mPendingCommands.Clear
mFutureWaiter.Cancel
mFutureBuilder.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractProcessor As ContractProcessor, _
                ByVal pContractStore As IContractStore, _
                ByVal pBatchOrders As Boolean, _
                ByVal pStageOrders As Boolean, _
                ByVal pMaxExpenditure As Long, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me

Set mContractSpec = pContractSpec
Set mContractProcessor = pContractProcessor
Set mContractStore = pContractStore
mBatchOrders = pBatchOrders
mStageOrders = pStageOrders
mMaxExpenditure = pMaxExpenditure
mUnderlyingExchangeName = pUnderlyingExchangeName

Set mFutureWaiter = New FutureWaiter

If pMaxExpenditure <> 0 Then
    gWriteLineToConsole "Resolving contract", True
    mFutureWaiter.Add GetOptionContract( _
                            pContractSpec, _
                            pContractStore, _
                            pMaxExpenditure, _
                            0, _
                            pUnderlyingExchangeName, _
                            pMarketDataManager, _
                            Me)
Else
    mFutureWaiter.Add FetchContract(pContractSpec, mContractStore)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processTargetContract(ByVal pContract As IContract)
Const ProcName As String = "processTargetContract"
On Error GoTo Err

mFutureBuilder.Value = pContract
mFutureBuilder.Complete
gWriteLineToConsole "Contract is: " & gGetContractName(pContract), True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processPendingCommands()
Const ProcName As String = "processPendingCommands"
On Error GoTo Err

Dim c As PendingCommand
For Each c In mPendingCommands
    gWriteLineToConsole c.Command.Name & " " & c.Params, True
    If c.Command Is gCommands.BracketCommand Then
        mContractProcessor.ProcessBracketCommand c.Params
    ElseIf c.Command Is gCommands.BuyCommand Then
        mContractProcessor.ProcessBuyCommand c.Params
    ElseIf c.Command Is gCommands.EndBracketCommand Then
        mContractProcessor.ProcessEndBracketCommand
    ElseIf c.Command Is gCommands.EntryCommand Then
        mContractProcessor.ProcessEntryCommand c.Params
    ElseIf c.Command Is gCommands.QuitCommand Then
        mContractProcessor.ProcessQuitCommand
    ElseIf c.Command Is gCommands.RolloverCommand Then
        mContractProcessor.ProcessRolloverCommand c.Params
    ElseIf c.Command Is gCommands.SellCommand Then
        mContractProcessor.ProcessSellCommand c.Params
    ElseIf c.Command Is gCommands.StopLossCommand Then
        mContractProcessor.ProcessStopLossCommand c.Params
    ElseIf c.Command Is gCommands.TargetCommand Then
        mContractProcessor.ProcessTargetCommand c.Params
    End If
Next

gPlaceOrdersTask.AddContractProcessor mContractProcessor, mStageOrders
mPendingCommands.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


