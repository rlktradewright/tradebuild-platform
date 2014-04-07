VERSION 5.00
Begin VB.Form fTickfileOrganiser 
   Caption         =   "Tickfile Organiser"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin TradingUI27.TickfileOrganiser TickfileOrganiser1 
      Height          =   4065
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   7170
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "fTickfileOrganiser"
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

Private Const ModuleName                    As String = "fTickfileOrganiser"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                          As Boolean

Private mMinimumWidth                       As Long
Private mMinimumHeight                      As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

mMinimumWidth = 120 + TickfileOrganiser1.MinimumWidth + 120 + OkButton.Width + 120
mMinimumHeight = 120 + TickfileOrganiser1.MinimumHeight + 120

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"
On Error GoTo Err

If mMinimumWidth = 0 Then Exit Sub

Me.ScaleMode = vbTwips
If Me.ScaleWidth < mMinimumWidth Then Me.ScaleWidth = mMinimumWidth
If Me.ScaleHeight < mMinimumHeight Then Me.ScaleHeight = mMinimumHeight

OkButton.Left = Me.ScaleWidth - OkButton.Width - 120
CancelButton.Left = OkButton.Left
TickfileOrganiser1.Width = OkButton.Left - 240

TickfileOrganiser1.Height = Me.ScaleHeight - 240

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"

On Error GoTo Err

mCancelled = True
Unload Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OKButton_Click()
Const ProcName As String = "OKButton_Click"

On Error GoTo Err

mCancelled = False
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickfileOrganiser1_TickfileCountChanged()
Const ProcName As String = "TickfileOrganiser1_TickfileCountChanged"
On Error GoTo Err

If TickfileOrganiser1.TickfileCount > 0 Then
    OkButton.Enabled = True
Else
    OkButton.Enabled = False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifiers
Const ProcName As String = "TickfileSpecifiers"
On Error GoTo Err

If Not mCancelled Then Set TickfileSpecifiers = TickfileOrganiser1.TickfileSpecifiers

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

TickfileOrganiser1.Initialise pTickfileStore, pPrimaryContractStore, pSecondaryContractStore

mCancelled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================





