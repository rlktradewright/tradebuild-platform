VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderRecoveryAgent"
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

Implements IOrderRecoveryAgent

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

Private Const ModuleName                            As String = "OrderRecoveryAgent"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderSubmissionSP                          As IOrderSubmissionSrvcProvider

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mOrderSubmissionSP = gTB.ServiceProviders.GetServiceProvider(LiveOrderSubmission)
Assert Not mOrderSubmissionSP Is Nothing, , "No live order submission service provider is configured"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IOrderRecoveryAgent Interface Members
'@================================================================================

Private Sub IOrderRecoveryAgent_FinishOrderRecovery()
mOrderSubmissionSP.FinishOrderRecovery
End Sub

Private Sub IOrderRecoveryAgent_StartOrderRecovery(ByVal pListener As IOrderRecoveryListener, ByVal pStartFrom As Date)
mOrderSubmissionSP.StartOrderRecovery pListener, pStartFrom
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




