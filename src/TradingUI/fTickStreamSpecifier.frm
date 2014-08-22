VERSION 5.00
Begin VB.Form fTickStreamSpecifier 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tickstream specifier"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TradingUI27.TickStreamSpecifier TickStreamSpecifier1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11853
      _ExtentY        =   7408
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "fTickStreamSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Const ModuleName                    As String = "fTickstreamSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                          As Boolean

Private mTickfileSpecifiers                 As TickfileSpecifiers

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Load()
mCancelled = True
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
mCancelled = True
Me.Hide
'Unload Me
End Sub

Private Sub OKButton_Click()
Const ProcName As String = "OKButton_Click"
On Error GoTo Err

Screen.MousePointer = vbHourglass
mCancelled = False
TickStreamSpecifier1.Load
TickStreamSpecifier1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickStreamSpecifier1_NotReady()
OkButton.Enabled = False
End Sub

Private Sub TickStreamSpecifier1_ready()
OkButton.Enabled = True
End Sub

Private Sub TickStreamSpecifier1_TickStreamsSpecified( _
                ByVal pTickfileSpecifiers As TickfileSpecifiers)
Screen.MousePointer = vbDefault
Set mTickfileSpecifiers = pTickfileSpecifiers
Me.Hide
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifiers
Set TickfileSpecifiers = mTickfileSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore)
Const ProcName As String = "Initialise"
On Error GoTo Err

TickStreamSpecifier1.Initialise pTickfileStore, pPrimaryContractStore, pSecondaryContractStore

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





