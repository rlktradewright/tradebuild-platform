VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#23.5#0"; "TWControls40.ocx"
Begin VB.Form fTimeframeSpecifier 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Specify a timeframe"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TradingUI27.TimeframeSpecifier TimeframeSpecifier1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin TWControls40.TWButton CancelButton 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   600
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
      Height          =   375
      Left            =   2760
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
      Caption         =   "Ok"
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

Private mCancelled As Boolean

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

'If TimeframeSpecifier1.isTimeframeValid Then
    OkButton.Enabled = True
'end if

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


