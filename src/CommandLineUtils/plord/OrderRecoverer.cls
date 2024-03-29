VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderRecoverer"
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

Private Const ModuleName                            As String = "OrderRecoverer"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTaskController                  As TaskController
Attribute mTaskController.VB_VarHelpID = -1
Private mOrderManager                               As OrderManager

Private mGroups                                     As Groups

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mTaskController Event Handlers
'@================================================================================

Private Sub mTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mTaskController_Completed"
On Error GoTo Err

gInputPaused = False
If ev.Cancelled Then
    gTerminate "Order recovery was cancelled - the program cannot continue."
    Exit Sub
End If

If ev.ErrorNumber <> 0 Then
    gTerminate "Order recovery encountered an error (" & ev.ErrorNumber & ") " & _
                        ev.ErrorMessage & vbCrLf & _
                        "At:" & vbCrLf & _
                        ev.ErrorSource & vbCrLf & _
                        ">>>>>>>                               <<<<<<<" & vbCrLf & _
                        ">>>>>>>   The program cannot continue <<<<<<<" & vbCrLf & _
                        ">>>>>>>                               <<<<<<<" & vbCrLf & _
                        vbCrLf & _
                        "You will need to delete the recovery file to enable the program to" & vbCrLf & _
                        "start properly. Make sure you manually manage any outstanding positions." & vbCrLf & _
                        vbCrLf
    Exit Sub
End If

LogMessage "Order recovery completed"

gCompleteOrderRecovery

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

Friend Sub RecoverOrders( _
                ByVal pOrderManager As OrderManager, _
                ByVal pScopeName As String, _
                ByVal pDataStore As IOrderPersistenceDataStore, _
                ByVal pRecoveryAgent As IOrderRecoveryAgent, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderSubmitterFactory As IOrderSubmitterFactory, _
                ByVal pGroups As Groups, _
                ByVal pMoneyManager As MoneyManager, _
                ByVal pAccountDataProvider As IAccountDataProvider, _
                ByVal pCurrencyConverter As ICurrencyConverter)
Const ProcName As String = "RecoverOrdersFromPreviousSession"
On Error GoTo Err

Set mOrderManager = pOrderManager
Set mGroups = pGroups
Set mTaskController = mOrderManager.RecoverOrdersFromPreviousSession( _
                                        pScopeName, _
                                        pDataStore, _
                                        pRecoveryAgent, _
                                        pMarketDataManager, _
                                        pOrderSubmitterFactory, _
                                        pMoneyManager, _
                                        pAccountDataProvider, _
                                        pCurrencyConverter)
gInputPaused = True
Set mSelfRef = Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




