VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractsBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
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

Private Const ModuleName                    As String = "ContractsBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mContracts                          As Contracts

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Contracts() As Contracts
Set Contracts = mContracts
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddContract(ByVal pContract As Contract)
Const ProcName As String = "AddContract"
Dim failpoint As String
On Error GoTo Err

mContracts.Add pContract

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub ContractSpecifierInvalid(ByVal reason As String)
Const ProcName As String = "ContractSpecifierInvalid"
Dim failpoint As String
On Error GoTo Err

mContracts.ContractSpecifierInvalid reason

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Friend Sub Initialise( _
                ByVal ContractSpec As ContractSpecifier)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mContracts = New Contracts
mContracts.Initialise ContractSpec

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub LoadComplete()
Const ProcName As String = "LoadComplete"
Dim failpoint As String
On Error GoTo Err

mContracts.LoadComplete

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub Notify( _
                ByRef ev As NotificationEvent)
Const ProcName As String = "Notify"
Dim failpoint As String
On Error GoTo Err

mContracts.Notify ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub Remove(ByVal index As Variant)
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

mContracts.Remove index

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub UnableToLoad( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)
Const ProcName As String = "UnableToLoad"
Dim failpoint As String
On Error GoTo Err

mContracts.ServiceProviderCantHandle reason, serviceProviderName

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


