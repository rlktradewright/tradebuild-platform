VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processor"
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

Private Const ModuleName                            As String = "Processor"

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                                 As IMarketDataSource

Private WithEvents mExitTLE                         As TimerListItem
Attribute mExitTLE.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Static sTickCount As Long
sTickCount = sTickCount + 1

If sTickCount = 1 Then
    gCon.WriteLineToConsole "Got first tick!"
Else
    gCon.WriteLineToConsole "Got another tick! - " & sTickCount
End If

mDataSource.RemoveGenericTickListener Me
mDataSource.Finish

Set mExitTLE = GetGlobalTimerList.Add(Nothing, 2, ExpiryTimeUnitSeconds)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mExitTLE Event Handlers
'@================================================================================

Private Sub mExitTLE_StateChange(ev As TWUtilities40.StateChangeEventData)
gFinish
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Process( _
    ByVal pContractSpec As IContractSpecifier, _
    ByVal pContractStore As IContractStore, _
    ByVal pMarketDataManager As RealTimeDataManager)
Const ProcName As String = "Process"
On Error GoTo Err

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




