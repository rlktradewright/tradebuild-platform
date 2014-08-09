VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimOrderSubmitterFactory"
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

Implements IOrderSubmitterFactory

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

Private Const ModuleName                            As String = "SimOrderSubmitterFactory"

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManager                          As IMarketDataManager
Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IOrderSubmitterFactory Interface Members
'@================================================================================

Private Function IOrderSubmitterFactory_CreateOrderSubmitter() As IOrderSubmitter
Const ProcName As String = "IOrderSubmitterFactory_CreateOrderSubmitter"
On Error GoTo Err

Dim lSubmitter As New SimulatedOrderSubmitter
lSubmitter.Initialise getOrderSimulator
Set IOrderSubmitterFactory_CreateOrderSubmitter = lSubmitter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore)
Set mMarketDataManager = pMarketDataManager
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getOrderSimulator() As InProcessExchangeSimulator
Const ProcName As String = "getOrderSimulator"
On Error GoTo Err

Static sSimulator As InProcessExchangeSimulator
If sSimulator Is Nothing Then
    Set sSimulator = New InProcessExchangeSimulator
    sSimulator.Initialise mMarketDataManager, mPrimaryContractStore, mSecondaryContractStore
End If
Set getOrderSimulator = sSimulator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




