VERSION 5.00
Object = "{D1E1CD3C-084A-4A4F-B2D9-56CE3669B04D}#1.0#0"; "TradeBuildUI.ocx"
Begin VB.Form fMarketDepth 
   Caption         =   "Market Depth"
   ClientHeight    =   3270
   ClientLeft      =   375
   ClientTop       =   510
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4335
   Begin TradeBuildUI.DOMDisplay DOMDisplay1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
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

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTicker As TradeBuild.Ticker
Attribute mTicker.VB_VarHelpID = -1
Private mCaption As String

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()

Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height

End Sub

Private Sub Form_Resize()
If Me.ScaleWidth = 0 And _
    Me.ScaleHeight = 0 Then Exit Sub
DOMDisplay1.Left = 0
DOMDisplay1.Top = 0
DOMDisplay1.Width = Me.ScaleWidth
DOMDisplay1.Height = Me.ScaleHeight
End Sub

Private Sub Form_Terminate()
Debug.Print "Market depth form terminated"
End Sub

Private Sub Form_Unload(cancel As Integer)
DOMDisplay1.finish
Set mTicker = Nothing
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub DOMDisplay1_Halted()
Me.caption = "Market depth data halted"
End Sub

Private Sub DOMDisplay1_Resumed()
Me.caption = mCaption
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_Error(ev As TradeBuild.ErrorEvent)
If ev.errorCode = ApiErrorCodes.ApiErrMarketDepthNotAvailable Then
    Unload Me
End If
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let numberOfRows(ByVal value As Long)
DOMDisplay1.numberOfRows = value
End Property

Public Property Let Ticker(ByVal value As Ticker)
Set mTicker = value
mCaption = "Market depth for " & _
            mTicker.Contract.specifier.localSymbol & _
            " on " & _
            mTicker.Contract.specifier.exchange
Me.caption = mCaption
DOMDisplay1.Ticker = mTicker
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

