VERSION 5.00
Object = "{41BEA792-C104-45F5-96C2-0BF81D749359}#1.0#0"; "TradeBuildUI.ocx"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4895
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

