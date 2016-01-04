VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
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
   Begin TWControls40.TWButton OkButton 
      Default         =   -1  'True
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   240
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      Appearance      =   0
      Caption         =   "Ok"
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
   End
   Begin TWControls40.TWButton CancelButton 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   840
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      Appearance      =   0
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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

Private Const ModuleName                    As String = "fTickstreamSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                          As Boolean

Private mTickfileSpecifiers                 As TickfileSpecifiers

Private mTheme                              As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Load()
mCancelled = True
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

Public Property Get Cancelled() As Boolean
Cancelled = mCancelled
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = value
If mTheme Is Nothing Then Exit Property

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





