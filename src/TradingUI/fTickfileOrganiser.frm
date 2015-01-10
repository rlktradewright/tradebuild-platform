VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#23.5#0"; "TWControls40.ocx"
Begin VB.Form fTickfileOrganiser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tickfile Organiser"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
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
   Begin TWControls40.TWButton CancelButton 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   720
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
   End
   Begin TWControls40.TWButton OkButton 
      Default         =   -1  'True
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Ok"
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

Implements IThemeable

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

Private mTheme                                  As ITheme

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
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

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

Public Property Get Cancelled() As Boolean
Cancelled = mCancelled
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = value
Me.BackColor = mTheme.BackColor
gApplyTheme mTheme, Me.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
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





