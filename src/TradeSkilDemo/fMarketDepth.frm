VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#307.0#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
Begin VB.Form fMarketDepth 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Market Depth"
   ClientHeight    =   5895
   ClientLeft      =   375
   ClientTop       =   390
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin TWControls40.TWButton CentreButton 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Centre"
      DefaultBorderColor=   15793920
      DisabledBackColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackColor=   0
      PushedBackColor =   0
   End
   Begin TradingUI27.DOMDisplay DOMDisplay1 
      Height          =   5520
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9737
   End
End
Attribute VB_Name = "fMarketDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements IErrorListener
Implements IThemeable

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                            As String = "fMarketDepth"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTicker                                     As Ticker
Attribute mTicker.VB_VarHelpID = -1
Private mCaption                                    As String

Private mTheme                                      As ITheme

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Load()

Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height

End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"

On Error GoTo Err

If Me.ScaleWidth = 0 And _
    Me.ScaleHeight = 0 Then Exit Sub

If Me.ScaleWidth / 2 - CentreButton.Width / 2 > 0 Then
    CentreButton.Left = Me.ScaleWidth / 2 - CentreButton.Width / 2
Else
    CentreButton.Left = 0
End If

DOMDisplay1.Width = Me.ScaleWidth
DOMDisplay1.Height = Me.ScaleHeight - DOMDisplay1.Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Terminate()
Debug.Print "Market depth form terminated"
End Sub

Private Sub Form_Unload(Cancel As Integer)
DOMDisplay1.Finish
Set mTicker = Nothing
End Sub

'================================================================================
' IErrorListener Interface Members
'================================================================================

Private Sub IErrorListener_Notify(ev As ErrorEventData)
Const ProcName As String = "IErrorListener_Notify"
On Error GoTo Err

gModelessMsgBox "Market depth is not available: " & ev.ErrorMessage, MsgBoxExclamation, mTheme, "Attention"
Unload Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub CentreButton_Click()
Const ProcName As String = "CentreButton_Click"
On Error GoTo Err

DOMDisplay1.Centre

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DOMDisplay1_Halted()
Const ProcName As String = "DOMDisplay1_Halted"
On Error GoTo Err

Me.caption = "Market depth data halted"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DOMDisplay1_Resumed()
Const ProcName As String = "DOMDisplay1_Resumed"
On Error GoTo Err

Me.caption = mCaption

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let NumberOfRows(ByVal Value As Long)
Const ProcName As String = "NumberOfRows"
On Error GoTo Err

DOMDisplay1.NumberOfRows = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
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

Public Property Let Ticker(ByVal Value As Ticker)
Const ProcName As String = "Ticker"
On Error GoTo Err

Set mTicker = Value
mTicker.AddErrorListener Me

Dim lContract As IContract
Set lContract = mTicker.ContractFuture.Value
mCaption = "Market depth for " & _
            lContract.Specifier.LocalSymbol & _
            " on " & _
            lContract.Specifier.Exchange
Me.caption = mCaption
DOMDisplay1.DataSource = mTicker

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


