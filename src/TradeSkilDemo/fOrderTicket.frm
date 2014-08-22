VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#235.0#0"; "TradingUI27.ocx"
Begin VB.Form fOrderTicket 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
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

Private mTicker                                 As Ticker

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Activate()
Const ProcName As String = "Form_Activate"
On Error GoTo Err

Me.left = CLng(mAppInstanceConfig.GetSetting(ConfigSettingOrderTicketLeft, 0)) * Screen.TwipsPerPixelX
Me.Top = CLng(mAppInstanceConfig.GetSetting(ConfigSettingOrderTicketTop, (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Deactivate()
Const ProcName As String = "Form_Deactivate"
On Error GoTo Err

updateSettings

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

Set mTicker = Nothing
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

Private Sub OrderTicket1_NeedLiveOrderContext()
Const ProcName As String = "OrderTicket1_NeedLiveOrderContext"
On Error GoTo Err

OrderTicket1.SetLiveOrderContext mTicker.PositionManager.OrderContexts.DefaultOrderContext

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrderTicket1_NeedSimulatedOrderContext()
Const ProcName As String = "OrderTicket1_NeedSimulatedOrderContext"
On Error GoTo Err

OrderTicket1.SetSimulatedOrderContext mTicker.PositionManagerSimulated.OrderContexts.DefaultOrderContext

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let Ticker(ByVal Value As Ticker)
Const ProcName As String = "Ticker"
On Error GoTo Err

If Value.State <> MarketDataSourceStateRunning Then Exit Property

If Not mTicker Is Nothing Then
    If mTicker Is Value Then Exit Property
End If

OrderTicket1.Clear

Set mTicker = Value

If mTicker.IsLiveOrdersEnabled And mTicker.IsSimulatedOrdersEnabled Then
    OrderTicket1.SetMode OrderTicketModeLiveAndSimulated
ElseIf mTicker.IsSimulatedOrdersEnabled Then
    OrderTicket1.SetMode OrderTicketModeSimulatedOnly
ElseIf mTicker.IsLiveOrdersEnabled Then
    OrderTicket1.SetMode OrderTicketModeLiveOnly
End If

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ShowBracketOrder( _
                ByVal Value As IBracketOrder, _
                ByVal selectedOrderNumber As Long)
Const ProcName As String = "ShowBracketOrder"
On Error GoTo Err

OrderTicket1.Clear
OrderTicket1.ShowBracketOrder Value, selectedOrderNumber

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

