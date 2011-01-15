VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickersStarter"
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

Event DisableCancelCommand()
Event DisableStartCommand()
Event EnableCancelCommand()
Event EnableStartCommand()
Event Error( _
                ByVal errorNumber, _
                ByVal errorMessage As String)
Event HideContractSelector()
Event NoContracts()
Event ShowContractSelector()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TickersStarter"

'@================================================================================
' Member variables
'@================================================================================

Private mWorkspace                                  As Workspace
Private mTickerOptions                              As TickerOptions
Private mContractSpec                               As contractSpecifier
Private mContractSelector                           As IContractSelector
Private mStartIntervalMillisecs                     As Long

Private WithEvents mContractsLoadTC                 As TaskController
Attribute mContractsLoadTC.VB_VarHelpID = -1
Private mContracts                                  As Contracts

Private WithEvents mStartTickersTC                  As TaskController
Attribute mStartTickersTC.VB_VarHelpID = -1
                
'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mContractsLoadTC_Completed"
Dim failpoint As String
On Error GoTo Err

If ev.errorNumber <> 0 Then
    RaiseEvent Error(ev.errorNumber, ev.errorMessage)
Else
    Set mContracts = ev.result
    handleContractsLoaded
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mStartTickersTC Event Handlers
'@================================================================================

Private Sub mStartTickersTC_Completed(ev As TWUtilities30.TaskCompletionEventData)
If ev.errorNumber <> 0 Then
    RaiseEvent Error(ev.errorNumber, ev.errorMessage)
Else
    RaiseEvent EnableStartCommand
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Cancel()
RaiseEvent HideContractSelector
RaiseEvent DisableCancelCommand
RaiseEvent DisableStartCommand
End Sub

Friend Sub Initialise( _
                ByVal pWorkspace As Workspace, _
                ByVal pTickerOptions As TickerOptions, _
                ByVal contractSpec As contractSpecifier, _
                ByVal pContractSelector As IContractSelector, _
                ByVal startIntervalMillisecs As Long)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mWorkspace = pWorkspace
mTickerOptions = pTickerOptions
Set mContractSpec = contractSpec
Set mContractSelector = pContractSelector
mStartIntervalMillisecs = startIntervalMillisecs

RaiseEvent DisableCancelCommand
RaiseEvent DisableStartCommand
RaiseEvent HideContractSelector

Set mContractsLoadTC = TradeBuildAPI.LoadContracts(mContractSpec)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Public Sub StartTickers()
Const ProcName As String = "StartTickers"
Dim failpoint As String
On Error GoTo Err

If mContractSelector.SelectedContracts.Count > 0 Then
    Set mStartTickersTC = mWorkspace.Tickers.StartTickersFromContracts( _
                                                mTickerOptions, _
                                                mContractSelector.SelectedContracts, _
                                                , _
                                                , _
                                                , _
                                                mStartIntervalMillisecs)
    RaiseEvent DisableStartCommand
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub handleContractsLoaded()
Const ProcName As String = "handleContractsLoaded"
Dim failpoint As String
On Error GoTo Err

If mContracts.Count = 0 Then
    RaiseEvent NoContracts
ElseIf mContracts.Count = 1 Then
    mWorkspace.Tickers.Add(mTickerOptions).StartTickerFromContract mContracts.item(1)
Else
    mContractSelector.Initialise mContracts
    RaiseEvent ShowContractSelector
    RaiseEvent EnableStartCommand
    RaiseEvent EnableCancelCommand
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


