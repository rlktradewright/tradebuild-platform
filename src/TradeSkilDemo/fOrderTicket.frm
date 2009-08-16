VERSION 5.00
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#166.3#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fOrderTicket 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin TradeBuildUI26.OrderTicket OrderTicket1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10821
   End
End
Attribute VB_Name = "fOrderTicket"
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

Me.left = CLng(gAppInstanceConfig.GetSetting(ConfigSettingOrderTicketLeft, 0)) * Screen.TwipsPerPixelX
Me.Top = CLng(gAppInstanceConfig.GetSetting(ConfigSettingOrderTicketTop, (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY

End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
updateSettings
End Sub

Private Sub Form_Unload(cancel As Integer)
OrderTicket1.Finish
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

Private Sub updateSettings()
If Not gAppInstanceConfig Is Nothing Then
    gAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionOrderTicket
    gAppInstanceConfig.SetSetting ConfigSettingOrderTicketLeft, Me.left / Screen.TwipsPerPixelX
    gAppInstanceConfig.SetSetting ConfigSettingOrderTicketTop, Me.Top / Screen.TwipsPerPixelY
End If
End Sub

