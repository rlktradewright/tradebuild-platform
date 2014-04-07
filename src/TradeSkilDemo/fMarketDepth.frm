VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#150.0#0"; "TradingUI27.ocx"
Begin VB.Form fMarketDepth 
   Caption         =   "Market Depth"
   ClientHeight    =   5895
   ClientLeft      =   375
   ClientTop       =   510
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   4335
   Begin VB.CommandButton CentreButton 
      Caption         =   "Centre"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1215
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

Implements IMarketDataErrorListener

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

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()

Me.left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height

End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"

On Error GoTo Err

If Me.ScaleWidth = 0 And _
    Me.ScaleHeight = 0 Then Exit Sub

If Me.ScaleWidth / 2 - CentreButton.Width / 2 > 0 Then
    CentreButton.left = Me.ScaleWidth / 2 - CentreButton.Width / 2
Else
    CentreButton.left = 0
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

Private Sub Form_Unload(cancel As Integer)
DOMDisplay1.Finish
Set mTicker = Nothing
End Sub

'================================================================================
' IMarketDataErrorListener Interface Members
'================================================================================

Private Sub IMarketDataErrorListener_NotifyMarketDataError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataErrorListener_NotifyMarketDataError"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataErrorListener_NotifyMarketDepthError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataErrorListener_NotifyMarketDepthError"
On Error GoTo Err

gModelessMsgBox "Market depth is not available: " & pErrorMessage, MsgBoxExclamation, "Attention"
Unload Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
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

Public Property Let numberOfRows(ByVal value As Long)
Const ProcName As String = "numberOfRows"
On Error GoTo Err

DOMDisplay1.numberOfRows = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Ticker(ByVal value As Ticker)
Const ProcName As String = "Ticker"
On Error GoTo Err

Set mTicker = value
mTicker.AddErrorListener Me

Dim lContract As IContract
Set lContract = mTicker.ContractFuture.value
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


