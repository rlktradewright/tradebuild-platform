VERSION 5.00
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#158.0#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fMarketDepth 
   Caption         =   "Market Depth"
   ClientHeight    =   3630
   ClientLeft      =   375
   ClientTop       =   510
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4335
   Begin VB.CommandButton CentreButton 
      Caption         =   "Centre"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin TradeBuildUI26.DOMDisplay DOMDisplay1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
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

Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1
Private mCaption As String

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
If Me.ScaleWidth = 0 And _
    Me.ScaleHeight = 0 Then Exit Sub

If Me.ScaleWidth / 2 - CentreButton.Width / 2 > 0 Then
    CentreButton.left = Me.ScaleWidth / 2 - CentreButton.Width / 2
Else
    CentreButton.left = 0
End If

DOMDisplay1.Width = Me.ScaleWidth
DOMDisplay1.Height = Me.ScaleHeight - DOMDisplay1.Top
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

Private Sub CentreButton_Click()
DOMDisplay1.centre
End Sub

Private Sub DOMDisplay1_Halted()
Me.caption = "Market depth data halted"
End Sub

Private Sub DOMDisplay1_Resumed()
Me.caption = mCaption
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_Notification(ev As NotificationEvent)
If ev.eventCode = ApiNotifyCodes.ApiNotifyMarketDepthNotAvailable Then
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
            mTicker.Contract.specifier.Exchange
Me.caption = mCaption
DOMDisplay1.Ticker = mTicker
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

