VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#375.0#0"; "TradingUI27.ocx"
Begin VB.Form fContractSelector 
   Caption         =   "Form2"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form2"
   ScaleHeight     =   6510
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin TradingUI27.ContractSearch ContractSearch1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11033
   End
End
Attribute VB_Name = "fContractSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Const ModuleName                            As String = "fContractSelector"

'@================================================================================
' Member variables
'@================================================================================

Private mContractStore                              As IContractStore
Private mAllowMultipleSelection                     As Boolean

Private mSelectedContracts                          As IContracts

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ContractSearch1_Action()
Const ProcName As String = "ContractSearch1_Action"
On Error GoTo Err

Set mSelectedContracts = ContractSearch1.SelectedContracts
Unload Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get SelectedContracts() As IContracts
Const ProcName As String = "SelectedContracts"
On Error GoTo Err

Set SelectedContracts = mSelectedContracts

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContracts As IContracts, _
                ByVal pContractStore As IContractStore, _
                ByVal pAllowMultipleSelection As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractStore = pContractStore
mAllowMultipleSelection = pAllowMultipleSelection

ContractSearch1.Initialise pContractStore, Nothing
ContractSearch1.AllowMultipleSelection = pAllowMultipleSelection
ContractSearch1.LoadContracts pContracts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




