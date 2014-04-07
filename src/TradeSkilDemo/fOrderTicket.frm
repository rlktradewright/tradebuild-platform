VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#188.0#0"; "TradingUI27.ocx"
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
   Begin TradingUI27.OrderTicket OrderTicket1 
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

Private Const ModuleName                        As String = "fOrderTicket"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mAppInstanceConfig                      As ConfigurationSection

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

updateSettings

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Unload(cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

OrderTicket1.Clear

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub OrderTicket1_CaptionChanged(ByVal caption As String)
Const ProcName As String = "OrderTicket1_CaptionChanged"
On Error GoTo Err

Me.caption = caption

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let Ticker(ByVal value As Ticker)
Const ProcName As String = "Ticker"
On Error GoTo Err

If value.IsTickReplay Then
    OrderTicket1.Clear
    Exit Property
End If

If value.State <> MarketDataSourceStateRunning Then Exit Property
OrderTicket1.SetOrderContexts value.PositionManager.OrderContexts.DefaultOrderContext, value.PositionManagerSimulated.OrderContexts.DefaultOrderContext

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise(ByVal pAppInstanceConfig As ConfigurationSection)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mAppInstanceConfig = pAppInstanceConfig
Me.left = CLng(mAppInstanceConfig.GetSetting(ConfigSettingOrderTicketLeft, 0)) * Screen.TwipsPerPixelX
Me.Top = CLng(mAppInstanceConfig.GetSetting(ConfigSettingOrderTicketTop, (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ShowBracketOrder( _
                ByVal value As IBracketOrder, _
                ByVal selectedOrderNumber As Long)
Const ProcName As String = "ShowBracketOrder"
On Error GoTo Err

OrderTicket1.ShowBracketOrder value, selectedOrderNumber

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub updateSettings()
Const ProcName As String = "updateSettings"

On Error GoTo Err

If Not mAppInstanceConfig Is Nothing Then
    mAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionOrderTicket
    mAppInstanceConfig.SetSetting ConfigSettingOrderTicketLeft, Me.left / Screen.TwipsPerPixelX
    mAppInstanceConfig.SetSetting ConfigSettingOrderTicketTop, Me.Top / Screen.TwipsPerPixelY
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

