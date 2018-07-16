VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.Form fTimeframeSpecifier 
   BorderStyle     =   0  'None
   Caption         =   "Specify a timeframe"
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TitleText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Specify a timeframe:"
      Top             =   120
      Width           =   2895
   End
   Begin TradingUI27.TimeframeSpecifier TimeframeSpecifier1 
      Height          =   855
      Left            =   150
      TabIndex        =   2
      Top             =   420
      Width           =   2415
      _ExtentX        =   4471
      _ExtentY        =   1508
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin TWControls40.TWButton CancelButton 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   900
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
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
   Begin TWControls40.TWButton OkButton 
      Default         =   -1  'True
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   420
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
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
   Begin VB.Shape FrameShape 
      BorderWidth     =   2
      Height          =   1410
      Left            =   0
      Top             =   0
      Width           =   3570
   End
End
Attribute VB_Name = "fTimeframeSpecifier"
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

Private Const ModuleName As String = "fTimeframeSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                          As Boolean

Private mTheme                              As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Activate()
Const ProcName As String = "Form_Activate"
On Error GoTo Err

mCancelled = False
TimeframeSpecifier1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

OkButton.Enabled = True

If Not mTheme Is Nothing Then applyTheme

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

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"

On Error GoTo Err

mCancelled = True
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OKButton_Click()
Const ProcName As String = "OKButton_Click"

On Error GoTo Err

Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' TimeframeSpecifier1 Event Handlers
'@================================================================================

Private Sub TimeframeSpecifier1_Change()
Const ProcName As String = "TimeframeSpecifier1_Change"

On Error GoTo Err

If TimeframeSpecifier1.IsTimeframeValid Then
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

Friend Property Get Cancelled() As Boolean
Cancelled = mCancelled
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

applyTheme

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Friend Property Get TimePeriod() As TimePeriod
Const ProcName As String = "TimePeriod"

On Error GoTo Err

Assert Not mCancelled, "Cancelled by user"

Set TimePeriod = TimeframeSpecifier1.TimePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pValidator As ITimePeriodValidator)
TimeframeSpecifier1.Initialise pValidator
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub applyTheme()
Const ProcName As String = "applyTheme"
On Error GoTo Err

Me.BackColor = mTheme.BackColor
Me.FrameShape.BorderColor = mTheme.AlertForeColor

gApplyTheme mTheme, Me.Controls

Me.TitleText.BackColor = mTheme.BackColor
Me.TitleText.ForeColor = mTheme.AlertForeColor

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
