VERSION 5.00
Object = "{0AA335E1-8A4A-48F4-ADFF-D2FBF7CBB54B}#2.0#0"; "TradeBuildUI.ocx"
Begin VB.Form OrderForm 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin TradeBuildUI.OrderTicket OrderTicket1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10821
   End
End
Attribute VB_Name = "OrderForm"
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
' Amendment history
'================================================================================
'
'
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

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()

Me.Left = 0
Me.Top = Screen.Height - Me.Height

End Sub

Private Sub Form_Terminate()
Debug.Print "OrderForm terminated"
End Sub

Private Sub Form_Unload(cancel As Integer)
Debug.Print "OrderForm unloaded"
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub OrderTicket1_CaptionChanged(ByVal caption As String)
Me.caption = caption
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let ordersAreSimulated(ByVal value As Boolean)
OrderTicket1.ordersAreSimulated = value
End Property

Public Property Let Ticker(ByVal value As Ticker)
OrderTicket1.Ticker = value
End Property

'================================================================================
' Methods
'================================================================================

Public Sub showOrderPlex( _
                ByVal value As OrderPlex, _
                ByVal selectedOrderNumber As Long)
OrderTicket1.showOrderPlex value, selectedOrderNumber
End Sub

'================================================================================
' Helper Functions
'================================================================================

