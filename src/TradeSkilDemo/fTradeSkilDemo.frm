VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#279.0#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#27.1#0"; "TWControls40.ocx"
Begin VB.Form fTradeSkilDemo 
   Caption         =   "TradeSkil Demo Edition"
   ClientHeight    =   9960
   ClientLeft      =   225
   ClientTop       =   345
   ClientWidth     =   16665
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   16665
   Begin VB.PictureBox ShowInfoPanelPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   16320
      MouseIcon       =   "fTradeSkilDemo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":0152
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      ToolTipText     =   "Show Information Panel"
      Top             =   9345
      Width           =   240
   End
   Begin VB.PictureBox HideInfoPanelPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   16290
      MouseIcon       =   "fTradeSkilDemo.frx":06DC
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":082E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      ToolTipText     =   "Hide Information Panel"
      Top             =   5070
      Width           =   240
   End
   Begin VB.PictureBox ShowFeaturesPanelPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      MouseIcon       =   "fTradeSkilDemo.frx":0DB8
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":0F0A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
      ToolTipText     =   "Show Features Panel"
      Top             =   120
      Width           =   240
   End
   Begin TabDlg.SSTab InfoSSTab 
      Height          =   4455
      Left            =   4320
      TabIndex        =   3
      Top             =   5040
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7858
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1. Orders"
      TabPicture(0)   =   "fTradeSkilDemo.frx":1494
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TickfileOrdersSummary"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SimulatedOrdersSummary"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LiveOrdersSummary"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OrdersSummaryTabStrip"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OrderTicket1Button"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ModifyOrderPlexButton"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CancelOrderPlexButton"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ClosePositionsButton"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&2. Executions"
      TabPicture(1)   =   "fTradeSkilDemo.frx":14B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TickfileExecutionsSummary"
      Tab(1).Control(1)=   "SimulatedExecutionsSummary"
      Tab(1).Control(2)=   "ExecutionsSummaryTabStrip"
      Tab(1).Control(3)=   "LiveExecutionsSummary"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&3. Log"
      TabPicture(2)   =   "fTradeSkilDemo.frx":14CC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LogText"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin TWControls40.TWButton ClosePositionsButton 
         Height          =   495
         Left            =   11160
         TabIndex        =   7
         Top             =   3480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DefaultBorderColor=   15793920
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Close all positions!"
      End
      Begin TWControls40.TWButton CancelOrderPlexButton 
         Height          =   495
         Left            =   11160
         TabIndex        =   6
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DefaultBorderColor=   15793920
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
         Caption         =   "&Cancel"
      End
      Begin TWControls40.TWButton ModifyOrderPlexButton 
         Height          =   495
         Left            =   11160
         TabIndex        =   5
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DefaultBorderColor=   15793920
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
         Caption         =   "&Modify"
      End
      Begin TWControls40.TWButton OrderTicket1Button 
         Height          =   495
         Left            =   11160
         TabIndex        =   4
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DefaultBorderColor=   15793920
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
         Caption         =   "Order Ticket"
      End
      Begin VB.TextBox LogText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Status messages"
         Top             =   120
         Width           =   11955
      End
      Begin MSComctlLib.TabStrip OrdersSummaryTabStrip 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         MultiRow        =   -1  'True
         Style           =   2
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Live"
               Object.ToolTipText     =   "Show live orders"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Simulated"
               Object.ToolTipText     =   "Show simulated orders"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tickfile"
               Object.ToolTipText     =   "Show tickfile orders"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TradingUI27.OrdersSummary LiveOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradingUI27.OrdersSummary SimulatedOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradingUI27.ExecutionsSummary LiveExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   13
         Top             =   120
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6376
      End
      Begin MSComctlLib.TabStrip ExecutionsSummaryTabStrip 
         Height          =   375
         Left            =   -74880
         TabIndex        =   14
         Top             =   3720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         MultiRow        =   -1  'True
         Style           =   2
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Live"
               Object.ToolTipText     =   "Show live executions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Simulated"
               Object.ToolTipText     =   "Show simulated executions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tickfile"
               Object.ToolTipText     =   "Show executions against tickfiles"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TradingUI27.ExecutionsSummary SimulatedExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   15
         Top             =   120
         Width           =   11995
         _ExtentX        =   21167
         _ExtentY        =   6376
      End
      Begin TradingUI27.OrdersSummary TickfileOrdersSummary 
         Height          =   3615
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6376
      End
      Begin TradingUI27.ExecutionsSummary TickfileExecutionsSummary 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   18
         Top             =   120
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   6376
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   9585
      Width           =   16665
      _ExtentX        =   29395
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   23733
            Key             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   "timezone"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Key             =   "datetime"
         EndProperty
      EndProperty
   End
   Begin TradingUI27.TickerGrid TickerGrid1 
      Height          =   4815
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8493
      AllowUserReordering=   3
      BackColorFixed  =   16053492
      RowBackColorOdd =   16316664
      RowBackColorEven=   15658734
      GridColorFixed  =   14737632
      ForeColorFixed  =   10526880
      ForeColor       =   7368816
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TradeSkilDemo27.FeaturesPanel FeaturesPanel 
      Height          =   9375
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   16536
   End
End
Attribute VB_Name = "fTradeSkilDemo"
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

Implements LogListener
Implements StateChangeListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================
    
Private Const ModuleName                    As String = "fTradeSkilDemo"

Private Const ExecutionsTabCaptionLive      As String = "Live"
Private Const ExecutionsTabCaptionSimulated As String = "Simulated"
Private Const ExecutionsTabCaptionTickfile  As String = "Tickfile"

'================================================================================
' Enums
'================================================================================

Private Enum InfoTabIndexNumbers
    InfoTabIndexOrders
    InfoTabIndexExecutions
    InfoTabIndexLog
End Enum

Private Enum OrdersTabIndexNumbers
    OrdersTabIndexLive = 1
    OrdersTabIndexSimulated
    OrderTabIndexTickfile
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI                   As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1

Private mTickers                                    As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private mFeaturesPanelHidden                        As Boolean
Private mFeaturesPanelPinned                        As Boolean

Private mInfoPanelHidden                            As Boolean

Private mClockDisplay                               As ClockDisplay

Private mAppInstanceConfig                          As ConfigurationSection

Private WithEvents mOrderRecoveryFutureWaiter       As FutureWaiter
Attribute mOrderRecoveryFutureWaiter.VB_VarHelpID = -1
Private WithEvents mContractsFutureWaiter           As FutureWaiter
Attribute mContractsFutureWaiter.VB_VarHelpID = -1

Private mChartForms                                 As New ChartForms

Private mPreviousMainForm                           As fTradeSkilDemo

Private mOrderTicket                                As fOrderTicket

Private WithEvents mFeaturesPanelForm               As fFeaturesPanel
Attribute mFeaturesPanelForm.VB_VarHelpID = -1

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitialiseCommonControls
Set mOrderRecoveryFutureWaiter = New FutureWaiter
Set mContractsFutureWaiter = New FutureWaiter
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

setupLogging

Set mClockDisplay = New ClockDisplay
mClockDisplay.Initialise StatusBar1.Panels("datetime"), StatusBar1.Panels("timezone")
mClockDisplay.SetClock getDefaultClock

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_QueryUnload( _
                Cancel As Integer, _
                UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

updateInstanceSettings

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"
On Error GoTo Err

Static prevHeight As Long
Static prevWidth As Long

If Me.WindowState = FormWindowStateConstants.vbMinimized Then Exit Sub

If Me.Width < FeaturesPanel.Width + 120 Then Me.Width = FeaturesPanel.Width + 120
If Me.Height < 9555 Then Me.Height = 9555

If Me.Width = prevWidth And Me.Height = prevHeight Then Exit Sub

prevWidth = Me.Width
prevHeight = Me.Height

Resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

LogMessage "Unloading main form"

LogMessage "Shutting down clock"
mClockDisplay.Finish

LogMessage "Finishing Features Panel"
FeaturesPanel.Finish
If Not mFeaturesPanelForm Is Nothing Then mFeaturesPanelForm.Finish

Shutdown

LogMessage "Closing charts and market depth forms"
closeChartsAndMarketDepthForms

LogMessage "Closing config editor form"
gUnloadConfigEditor

LogMessage "Closing other forms"
Dim f As Form
For Each f In Forms
    If Not TypeOf f Is fTradeSkilDemo And Not TypeOf f Is fSplash Then
        LogMessage "Closing form: caption=" & f.caption & "; type=" & TypeName(f)
        Unload f
    End If
Next

LogMessage "Stopping tickers"
If Not mTickers Is Nothing Then mTickers.Finish

killLoggingForThisForm

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' LogListener Interface Members
'================================================================================

Private Sub LogListener_Finish()
'nothing to do
End Sub

Private Sub LogListener_Notify(ByVal Logrec As LogRecord)
Const ProcName As String = "LogListener_Notify"
On Error GoTo Err

If Len(LogText.Text) >= 32767 Then
    ' clear some space at the start of the textbox
    LogText.SelStart = 0
    LogText.SelLength = 16384
    LogText.SelText = ""
End If

LogText.SelStart = Len(LogText.Text)
LogText.SelLength = 0
If Len(LogText.Text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = formatLogRecord(Logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

OrderTicket1Button.Enabled = Not (getSelectedDataSource Is Nothing)
Dim lDataSource As IMarketDataSource
Set lDataSource = ev.Source

Select Case ev.State
Case MarketDataSourceStates.MarketDataSourceStateCreated

Case MarketDataSourceStates.MarketDataSourceStateReady
    If lDataSource Is getSelectedDataSource Then mClockDisplay.SetClockFuture lDataSource.ClockFuture
Case MarketDataSourceStates.MarketDataSourceStateRunning
    
Case MarketDataSourceStates.MarketDataSourceStatePaused

Case MarketDataSourceStates.MarketDataSourceStateStopped
    If getSelectedDataSource Is Nothing Then
    Else
        mClockDisplay.SetClockFuture getSelectedDataSource.ClockFuture
    End If
    
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub CancelOrderPlexButton_Click()
Const ProcName As String = "CancelOrderPlexButton_Click"
On Error GoTo Err

Dim op As IBracketOrder

If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
    Set op = LiveOrdersSummary.SelectedItem
ElseIf OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexSimulated Then
    Set op = SimulatedOrdersSummary.SelectedItem
Else
    Set op = TickfileOrdersSummary.SelectedItem
End If
If Not op Is Nothing Then op.Cancel True

CancelOrderPlexButton.Enabled = False
ModifyOrderPlexButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ClosePositionsButton_Click()
Const ProcName As String = "ClosePositionsButton_Click"
On Error GoTo Err

If Not mTradeBuildAPI.ClosingPositions Then
    If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
        mTradeBuildAPI.CloseAllPositions PositionTypeLive, _
                                        ClosePositionCancelOrders Or ClosePositionWaitForCancel
    Else
        mTradeBuildAPI.CloseAllPositions PositionTypeSimulated, _
                                        ClosePositionCancelOrders Or ClosePositionWaitForCancel
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ExecutionsSummaryTabStrip_Click()
Const ProcName As String = "ExecutionsSummaryTabStrip_Click"
On Error GoTo Err

Static currIndex As Long
If ExecutionsSummaryTabStrip.SelectedItem.Index = currIndex Then Exit Sub

Select Case ExecutionsSummaryTabStrip.SelectedItem.caption
Case ExecutionsTabCaptionLive
    LiveExecutionsSummary.Visible = True
    SimulatedExecutionsSummary.Visible = False
    TickfileExecutionsSummary.Visible = False
Case ExecutionsTabCaptionSimulated
    LiveExecutionsSummary.Visible = False
    SimulatedExecutionsSummary.Visible = True
    TickfileExecutionsSummary.Visible = False
Case ExecutionsTabCaptionTickfile
    LiveExecutionsSummary.Visible = False
    SimulatedExecutionsSummary.Visible = False
    TickfileExecutionsSummary.Visible = True
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_Hide()
Const ProcName As String = "FeaturesPanel_Hide"
On Error GoTo Err

FeaturesPanel.Visible = False
mFeaturesPanelHidden = True
Resize
updateInstanceSettings
ShowFeaturesPanelPicture.Visible = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_Unpin()
Const ProcName As String = "FeaturesPanel_Unpin"
On Error GoTo Err

FeaturesPanel.Visible = False
mFeaturesPanelPinned = False
Resize
updateInstanceSettings

mFeaturesPanelForm.Show vbModeless, Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub HideInfoPanelPicture_Click()
Const ProcName As String = "HideInfoPanelPicture_Click"
On Error GoTo Err

hideInfoPanel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub InfoSSTab_Click(PreviousTab As Integer)
Const ProcName As String = "InfoSSTAB_Click"
On Error GoTo Err

Select Case InfoSSTab.Tab
Case InfoSSTab.Tab = InfoTabIndexNumbers.InfoTabIndexLog
Case InfoSSTab.Tab = InfoTabIndexNumbers.InfoTabIndexOrders
    If ModifyOrderPlexButton.Enabled Then
        ModifyOrderPlexButton.Default = True
    Else
        If CancelOrderPlexButton.Enabled Then CancelOrderPlexButton.Default = True
    End If
Case InfoSSTab.Tab = InfoTabIndexNumbers.InfoTabIndexExecutions
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LiveOrdersSummary_SelectionChanged()
Const ProcName As String = "LiveOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection LiveOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ModifyOrderPlexButton_Click()
Const ProcName As String = "ModifyOrderPlexButton_Click"
On Error GoTo Err

Dim os As OrdersSummary

If OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexLive Then
    Set os = LiveOrdersSummary
ElseIf OrdersSummaryTabStrip.SelectedItem.Index = OrdersTabIndexNumbers.OrdersTabIndexSimulated Then
    Set os = SimulatedOrdersSummary
Else
    Set os = TickfileOrdersSummary
End If

If os.SelectedItem Is Nothing Then
    ModifyOrderPlexButton.Enabled = False
ElseIf os.IsSelectedItemModifiable Then
    mOrderTicket.Show vbModeless, Me
    mOrderTicket.ShowBracketOrder os.SelectedItem, os.SelectedOrderIndex
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrdersSummaryTabStrip_Click()
Const ProcName As String = "OrdersSummaryTabStrip_Click"
On Error GoTo Err

Static currIndex As Long
If OrdersSummaryTabStrip.SelectedItem.Index = currIndex Then Exit Sub

Select Case OrdersSummaryTabStrip.SelectedItem.Index
Case OrdersTabIndexNumbers.OrdersTabIndexLive
    LiveOrdersSummary.Visible = True
    SimulatedOrdersSummary.Visible = False
    TickfileOrdersSummary.Visible = False
    setOrdersSelection LiveOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrdersTabIndexLive
Case OrdersTabIndexNumbers.OrdersTabIndexSimulated
    LiveOrdersSummary.Visible = False
    SimulatedOrdersSummary.Visible = True
    TickfileOrdersSummary.Visible = False
    setOrdersSelection SimulatedOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrdersTabIndexSimulated
Case OrdersTabIndexNumbers.OrderTabIndexTickfile
    LiveOrdersSummary.Visible = False
    SimulatedOrdersSummary.Visible = False
    TickfileOrdersSummary.Visible = True
    setOrdersSelection TickfileOrdersSummary
    currIndex = OrdersTabIndexNumbers.OrderTabIndexTickfile
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OrderTicket1Button_Click()
Const ProcName As String = "OrderTicket1Button_Click"
On Error GoTo Err

showOrderTicket

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ShowFeaturesPanelPicture_Click()
Const ProcName As String = "ShowFeaturesPanelPicture_Click"
On Error GoTo Err

showFeaturesPanel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

Private Sub ShowInfoPanelPicture_Click()
Const ProcName As String = "ShowInfoPanelPicture_Click"
On Error GoTo Err

showInfoPanel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub SimulatedOrdersSummary_SelectionChanged()
Const ProcName As String = "SimulatedOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection SimulatedOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TickerGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "TickerGrid1_KeyUp"
On Error GoTo Err

Select Case KeyCode
Case vbKeyDelete
    StopSelectedTickers
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TickerGrid1_TickerSelectionChanged()
Const ProcName As String = "TickerGrid1_TickerSelectionChanged"
On Error GoTo Err

handleSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TickerGrid1_TickerSymbolEntered(ByVal pSymbol As String, ByVal pPreferredRow As Long)
Const ProcName As String = "TickerGrid1_TickerSymbolEntered"
On Error GoTo Err

mContractsFutureWaiter.Add FetchContracts(CreateContractSpecifier(, pSymbol), mTradeBuildAPI.ContractStorePrimary, mTradeBuildAPI.ContractStoreSecondary), pPreferredRow

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickfileOrdersSummary_SelectionChanged()
Const ProcName As String = "SimulatedOrdersSummary_SelectionChanged"
On Error GoTo Err

setOrdersSelection TickfileOrdersSummary

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mContractsFutureWaiter Event Handlers
'================================================================================

Private Sub mContractsFutureWaiter_WaitCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
Const ProcName As String = "mContractsFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Dim lContracts As IContracts
Set lContracts = ev.Future.Value

If lContracts.Count = 1 Then
    If IsContractExpired(lContracts.ItemAtIndex(1)) Then
        gModelessMsgBox "Contract has expired", MsgBoxExclamation, "Attention"
    Else
        TickerGrid1.StartTickerFromContract lContracts.ItemAtIndex(1), CLng(ev.ContinuationData)
    End If
Else
    If mFeaturesPanelHidden Then showFeaturesPanel
    FeaturesPanel.ShowTickersPane
    FeaturesPanel.LoadContractsForUserChoice lContracts, CLng(ev.ContinuationData)
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mFeaturesPanelForm Event Handlers
'================================================================================

Private Sub mFeaturesPanelForm_Hide()
Const ProcName As String = "mFeaturesPanelForm_Hide"
On Error GoTo Err

mFeaturesPanelForm.Hide
mFeaturesPanelHidden = True
updateInstanceSettings
ShowFeaturesPanelPicture.Visible = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_Pin()
Const ProcName As String = "mFeaturesPanelForm_Pin"
On Error GoTo Err

mFeaturesPanelForm.Hide
mFeaturesPanelPinned = True
FeaturesPanel.Visible = True
Resize
updateInstanceSettings

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mOrderRecoveryFutureWaiter Event Handlers
'================================================================================

Private Sub mOrderRecoveryFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mOrderRecoveryFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    LogMessage "Order recovery failed"
ElseIf ev.Future.IsAvailable Then
    LogMessage "Order recovery completed    "
    loadAppInstanceConfig
    
    Me.Show vbModeless
    If Not mFeaturesPanelPinned And Not mFeaturesPanelHidden Then mFeaturesPanelForm.Show vbModeless, Me

    gUnloadSplashScreen
    
    If Not mPreviousMainForm Is Nothing Then
        Unload mPreviousMainForm
        Set mPreviousMainForm = Nothing
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_Notification( _
                ByRef ev As NotificationEventData)
Const ProcName As String = "mTradeBuildAPI_Notification"
On Error GoTo Err

Select Case ev.EventCode
Case ApiNotifyCodes.ApiNotifyServiceProviderError
    Dim spError As ServiceProviderError
    Set spError = mTradeBuildAPI.GetServiceProviderError
    LogMessage "Error from " & _
                        spError.ServiceProviderName & _
                        ": code " & spError.ErrorCode & _
                        ": " & spError.Message

Case Else
    LogMessage "Notification: code=" & ev.EventCode & "; source=" & TypeName(ev.Source) & ": " & _
                ev.EventMessage & vbCrLf
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Function Initialise( _
                ByVal pTradeBuildAPI As TradeBuildAPI, _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                ByRef pErrorMessage As String) As Boolean
Const ProcName As String = "initialise"
On Error GoTo Err

Set mTradeBuildAPI = pTradeBuildAPI
Set mConfigStore = pConfigStore
Set mAppInstanceConfig = pAppInstanceConfig
Set mPreviousMainForm = gMainForm

LogMessage "Loading configuration: " & mAppInstanceConfig.InstanceQualifier

mAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionApplication

Set mTickers = mTradeBuildAPI.Tickers
If mTickers Is Nothing Then
    pErrorMessage = "No tickers object is available: one or more service providers may be missing or disabled"
    Initialise = False
    Exit Function
End If

LogMessage "Recovering orders from last session"
mOrderRecoveryFutureWaiter.Add CreateFutureFromTask(mTradeBuildAPI.RecoverOrders())

Initialise = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Shutdown()
Const ProcName As String = "Shutdown"
On Error GoTo Err

Static sAlreadyShutdown As Boolean
If sAlreadyShutdown Then Exit Sub

sAlreadyShutdown = True

LogMessage "Finishing UI controls"
finishUIControls

LogMessage "Removing service providers"
mTradeBuildAPI.ServiceProviders.RemoveAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyInstanceSettings()
Const ProcName As String = "applyInstanceSettings"
On Error GoTo Err

LogMessage "Loading configuration: positioning main form"
Select Case mAppInstanceConfig.GetSetting(ConfigSettingMainFormWindowstate, WindowStateNormal)
Case WindowStateMaximized
    Me.WindowState = FormWindowStateConstants.vbMaximized
Case WindowStateMinimized
    Me.WindowState = FormWindowStateConstants.vbMinimized
Case WindowStateNormal
    Me.left = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormLeft, 0)) * Screen.TwipsPerPixelX
    Me.Top = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormTop, 0)) * Screen.TwipsPerPixelY
    Me.Width = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormWidth, Me.Width / Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
    Me.Height = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormHeight, Me.Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
End Select

mInfoPanelHidden = CBool(mAppInstanceConfig.GetSetting(ConfigSettingMainFormFeaturesHidden, CStr(False)))
If mInfoPanelHidden Then
    hideInfoPanel
Else
    showInfoPanel
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub applyTheme(ByVal pTheme As ITheme)
Const ProcName As String = "applyTheme"
On Error GoTo Err

gTheme = pTheme

Me.BackColor = gTheme.BaseColor
gApplyTheme gTheme, Me.Controls

Dim lForm As Object
For Each lForm In Forms
    If TypeOf lForm Is IThemeable Then lForm.Theme = gTheme
Next

mChartForms.Theme = gTheme

SendMessage StatusBar1.hWnd, SB_SETBKCOLOR, 0, NormalizeColor(gTheme.StatusBarBackColor)

Dim lhDC As Long
lhDC = GetDC(StatusBar1.hWnd)
SetTextColor lhDC, NormalizeColor(gTheme.StatusBarForeColor)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearSelectedTickers()
Const ProcName As String = "clearSelectedTickers"
On Error GoTo Err

TickerGrid1.DeselectSelectedTickers
handleSelectedTickers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub closeChartsAndMarketDepthForms()
Const ProcName As String = "closeChartsAndMarketDepthForms"
On Error GoTo Err

mChartForms.Finish

Dim f As Form
For Each f In Forms
    If TypeOf f Is fMarketDepth Then
        LogMessage "Closing form: caption=" & f.caption & "; type=" & TypeName(f)
        Unload f
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub finishUIControls()
Const ProcName As String = "finishUIControls"
On Error GoTo Err

LiveOrdersSummary.Finish
SimulatedOrdersSummary.Finish
LiveExecutionsSummary.Finish
SimulatedExecutionsSummary.Finish
TickerGrid1.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function formatLogRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "formatLogRecord"
On Error GoTo Err

Static formatter As LogFormatter
If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.FormatRecord(Logrec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getDefaultClock() As Clock
Const ProcName As String = "getDefaultClock"
On Error GoTo Err

Static sClock As Clock
If sClock Is Nothing Then Set sClock = GetClock("") ' create a clock running local time
Set getDefaultClock = sClock

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSelectedDataSource() As IMarketDataSource
Const ProcName As String = "getSelectedDataSource"
On Error GoTo Err

If TickerGrid1.SelectedTickers.Count = 1 Then Set getSelectedDataSource = TickerGrid1.SelectedTickers.Item(1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleSelectedTickers()
Const ProcName As String = "handleSelectedTickers"
On Error GoTo Err

If TickerGrid1.SelectedTickers.Count = 0 Then
    OrderTicket1Button.Enabled = False
    mClockDisplay.SetClock getDefaultClock
Else
    OrderTicket1Button.Enabled = False
    
    Dim lTicker As Ticker
    Set lTicker = getSelectedDataSource
    If lTicker Is Nothing Then
        mClockDisplay.SetClock getDefaultClock
    ElseIf lTicker.State = MarketDataSourceStateRunning Then
        mClockDisplay.SetClockFuture lTicker.ClockFuture
        Dim lContract As IContract
        Set lContract = lTicker.ContractFuture.Value
        If (lTicker.IsLiveOrdersEnabled Or lTicker.IsSimulatedOrdersEnabled) And lContract.Specifier.SecType <> SecTypeIndex Then
            OrderTicket1Button.Enabled = True
        End If
    Else
        mClockDisplay.SetClock getDefaultClock
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub hideFeaturesPanel()
Const ProcName As String = "hideFeaturesPanel"
On Error GoTo Err

FeaturesPanel.Visible = True
Resize
Me.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub hideInfoPanel()
Const ProcName As String = "hideInfoPanel"
On Error GoTo Err

InfoSSTab.Visible = False
ShowInfoPanelPicture.Visible = True
HideInfoPanelPicture.Visible = False
mInfoPanelHidden = True
Resize
Me.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub killLoggingForThisForm()
Const ProcName As String = "killLoggingForThisForm"
On Error GoTo Err

GetLogger("log").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadAppInstanceConfig()
Const ProcName As String = "loadAppInstanceConfig"
On Error GoTo Err

LogMessage "Loading configuration: " & mAppInstanceConfig.InstanceQualifier

LogMessage "Loading configuration: Setting up ticker grid"
setupTickerGrid

LogMessage "Loading configuration: Setting up order summaries"
setupOrderSummaries

LogMessage "Loading configuration: Setting up execution summaries"
setupExecutionSummaries

LogMessage "Loading configuration: setting up order ticket"
setupOrderTicket

applyInstanceSettings

LogMessage "Loading configuration: loading tickers into ticker grid"
TickerGrid1.LoadFromConfig mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionTickerGrid)

LogMessage "Loading configuration: loading default study configurations"
LoadDefaultStudyConfigurationsFromConfig mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionDefaultStudyConfigs)

LogMessage "Loading configuration: creating charts"
startCharts

LogMessage "Loading configuration: creating historical charts"
startHistoricalCharts

InfoSSTab.Tab = InfoTabIndexNumbers.InfoTabIndexOrders

LogMessage "Loading configuration: initialising Features Panels"
setupFeaturesPanels

LogMessage "Loading configuration: applying theme"
applyTheme New BlackTheme

Me.caption = gAppTitle & _
            " - " & mAppInstanceConfig.InstanceQualifier

LogMessage "Loaded configuration: " & mAppInstanceConfig.InstanceQualifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Resize()
Const ProcName As String = "Resize"
On Error GoTo Err

Dim left As Long
If mFeaturesPanelHidden Or Not mFeaturesPanelPinned Then
    left = ShowFeaturesPanelPicture.Width + 60
Else
    left = 120 + FeaturesPanel.Width + 120
End If

'StatusBar1.Top = Me.ScaleHeight - StatusBar1.Height

FeaturesPanel.Height = StatusBar1.Top - FeaturesPanel.Top

InfoSSTab.Move left, _
                    StatusBar1.Top - InfoSSTab.Height, _
                    Me.ScaleWidth - left - 120

HideInfoPanelPicture.Move InfoSSTab.left + InfoSSTab.Width - HideInfoPanelPicture.Width - 2 * Screen.TwipsPerPixelX, _
                        InfoSSTab.Top + Screen.TwipsPerPixelY
ShowInfoPanelPicture.Move HideInfoPanelPicture.left, _
                        StatusBar1.Top - 240

TickerGrid1.Move left, _
                TickerGrid1.Top, _
                Me.ScaleWidth - left - 120, _
                IIf(mInfoPanelHidden, ShowInfoPanelPicture.Top - 60, InfoSSTab.Top - 120) - TickerGrid1.Top

If OrderTicket1Button.left >= 0 Then
    OrderTicket1Button.left = InfoSSTab.Width - OrderTicket1Button.Width - 120
    ModifyOrderPlexButton.left = InfoSSTab.Width - ModifyOrderPlexButton.Width - 120
    CancelOrderPlexButton.left = InfoSSTab.Width - CancelOrderPlexButton.Width - 120
    ClosePositionsButton.left = InfoSSTab.Width - CancelOrderPlexButton.Width - 120
    
    LiveOrdersSummary.Width = ModifyOrderPlexButton.left - 120 - 120
    SimulatedOrdersSummary.Width = LiveOrdersSummary.Width
    TickfileOrdersSummary.Width = LiveOrdersSummary.Width
Else
    OrderTicket1Button.left = InfoSSTab.Width - OrderTicket1Button.Width - 120 - SSTabInactiveControlAdjustment
    ModifyOrderPlexButton.left = InfoSSTab.Width - ModifyOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    CancelOrderPlexButton.left = InfoSSTab.Width - CancelOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    ClosePositionsButton.left = InfoSSTab.Width - CancelOrderPlexButton.Width - 120 - SSTabInactiveControlAdjustment
    
    LiveOrdersSummary.Width = ModifyOrderPlexButton.left + SSTabInactiveControlAdjustment - 120 - 120
    SimulatedOrdersSummary.Width = LiveOrdersSummary.Width
    TickfileOrdersSummary.Width = LiveOrdersSummary.Width
End If

LogText.Width = InfoSSTab.Width - 120 - 120
LiveExecutionsSummary.Width = InfoSSTab.Width - 120 - 120
SimulatedExecutionsSummary.Width = InfoSSTab.Width - 120 - 120

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrdersSelection( _
                ByVal pOrdersSummary As OrdersSummary)
Const ProcName As String = "setOrdersSelection"
On Error GoTo Err

If pOrdersSummary.IsEditing Then
    pOrdersSummary.Default = True
    Exit Sub
End If

pOrdersSummary.Default = False

Dim selection As IBracketOrder
Set selection = pOrdersSummary.SelectedItem

If selection Is Nothing Then
    CancelOrderPlexButton.Enabled = False
    ModifyOrderPlexButton.Enabled = False
Else
    If pOrdersSummary.SelectedOrderIndex = 0 Then
        CancelOrderPlexButton.Enabled = True
    Else
        CancelOrderPlexButton.Enabled = False
    End If
    If pOrdersSummary.IsSelectedItemModifiable Then
        ModifyOrderPlexButton.Enabled = True
        ModifyOrderPlexButton.Default = True
    Else
        ModifyOrderPlexButton.Enabled = False
        ModifyOrderPlexButton.Default = False
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupExecutionSummaries()
Const ProcName As String = "setupExecutionSummaries"
On Error GoTo Err

Do While ExecutionsSummaryTabStrip.Tabs.Count > 0
    ExecutionsSummaryTabStrip.Tabs.Remove 1
Loop

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    SimulatedExecutionsSummary.Visible = True
    ExecutionsSummaryTabStrip.Tabs.Add 1, , ExecutionsTabCaptionSimulated
Else
    SimulatedExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersSimulated
    SimulatedExecutionsSummary.Visible = False
    LiveExecutionsSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    LiveExecutionsSummary.Visible = True
    ExecutionsSummaryTabStrip.Tabs.Add 1, , ExecutionsTabCaptionLive
    ExecutionsSummaryTabStrip.Tabs.Add 2, , ExecutionsTabCaptionSimulated
End If

If Not mTradeBuildAPI.TickfileStoreInput Is Nothing Then
    TickfileExecutionsSummary.Visible = False
    ExecutionsSummaryTabStrip.Tabs.Add ExecutionsSummaryTabStrip.Tabs.Count + 1, , ExecutionsTabCaptionTickfile
End If

If ExecutionsSummaryTabStrip.Tabs.Count = 1 Then
    ExecutionsSummaryTabStrip.Visible = False
    SimulatedExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top + ExecutionsSummaryTabStrip.Height - SimulatedExecutionsSummary.Top
Else
    ExecutionsSummaryTabStrip.Visible = True
    SimulatedExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
    LiveExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
    TickfileExecutionsSummary.Height = ExecutionsSummaryTabStrip.Top - SimulatedExecutionsSummary.Top
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupFeaturesPanels()
Const ProcName As String = "setupFeaturesPanels"
On Error GoTo Err

FeaturesPanel.Initialise True, mTradeBuildAPI, mAppInstanceConfig, TickerGrid1, TickfileOrdersSummary, TickfileExecutionsSummary, mChartForms, mOrderTicket
mFeaturesPanelPinned = CBool(mAppInstanceConfig.GetSetting(ConfigSettingFeaturesPanelPinned, "True"))
mFeaturesPanelHidden = CBool(mAppInstanceConfig.GetSetting(ConfigSettingFeaturesPanelHidden, "False"))
    
'FeaturesPanel.Visible = mFeaturesPanelPinned And Not mFeaturesPanelHidden
'ShowFeaturesPanelPicture.Visible = mFeaturesPanelHidden

Set mFeaturesPanelForm = New fFeaturesPanel
mFeaturesPanelForm.Initialise False, mTradeBuildAPI, mAppInstanceConfig, TickerGrid1, TickfileOrdersSummary, TickfileExecutionsSummary, mChartForms, mOrderTicket
'mFeaturesPanelForm.Theme = gTheme

If Not mFeaturesPanelHidden Then showFeaturesPanel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupLogging()
Const ProcName As String = "setupLogging"
On Error GoTo Err

GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupOrderSummaries()
Const ProcName As String = "setupOrderSummaries"
On Error GoTo Err

If mTradeBuildAPI.AllOrdersSimulated Then
    SimulatedOrdersSummary.Height = OrdersSummaryTabStrip.Top + OrdersSummaryTabStrip.Height - SimulatedOrdersSummary.Top
    SimulatedOrdersSummary.Visible = True
    
    LiveOrdersSummary.Visible = False
    
    OrdersSummaryTabStrip.Visible = False
    OrdersSummaryTabStrip.Tabs.Item(OrdersTabIndexSimulated).Selected = True
Else
    SimulatedOrdersSummary.Height = OrdersSummaryTabStrip.Top - SimulatedOrdersSummary.Top
    
    LiveOrdersSummary.Initialise mTradeBuildAPI.MarketDataManager
    LiveOrdersSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersLive
    LiveOrdersSummary.Height = SimulatedOrdersSummary.Height
    
    OrdersSummaryTabStrip.Visible = True
    OrdersSummaryTabStrip.Tabs.Item(OrdersTabIndexLive).Selected = True
End If

SimulatedOrdersSummary.Initialise mTradeBuildAPI.MarketDataManager
SimulatedOrdersSummary.MonitorPositions mTradeBuildAPI.OrderManager.PositionManagersSimulated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupOrderTicket()
Const ProcName As String = "setupOrderTicket"
On Error GoTo Err

Set mOrderTicket = New fOrderTicket
mOrderTicket.Initialise mAppInstanceConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTickerGrid()
Const ProcName As String = "setupTickerGrid"
On Error GoTo Err

TickerGrid1.Initialise mTickers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showFeaturesPanel()
Const ProcName As String = "showFeaturesPanel"
On Error GoTo Err

mFeaturesPanelHidden = False
If mFeaturesPanelPinned Then
    FeaturesPanel.Visible = True
    Resize
    Me.Refresh
Else
    mFeaturesPanelForm.Show vbModeless, Me
End If
updateInstanceSettings

ShowFeaturesPanelPicture.Visible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showInfoPanel()
Const ProcName As String = "showInfoPanel"
On Error GoTo Err

mInfoPanelHidden = False
Resize
Me.Refresh
InfoSSTab.Visible = True
ShowInfoPanelPicture.Visible = False
HideInfoPanelPicture.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showOrderTicket()
Const ProcName As String = "showOrderTicket"
On Error GoTo Err

If getSelectedDataSource Is Nothing Then
    gModelessMsgBox "No ticker selected - please select a ticker", vbExclamation, "Error"
Else
    mOrderTicket.Show vbModeless, Me
    mOrderTicket.Ticker = getSelectedDataSource
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub startCharts()
Const ProcName As String = "startCharts"
On Error GoTo Err

mChartForms.LoadChartsFromConfig mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionCharts), _
                                mTickers, _
                                mTradeBuildAPI.BarFormatterLibManager, _
                                mTradeBuildAPI.HistoricalDataStoreInput.TimePeriodValidator, _
                                gMainForm

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub startHistoricalCharts()
Const ProcName As String = "startHistoricalCharts"
On Error GoTo Err

mChartForms.LoadHistoricalChartsFromConfig _
                    mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionHistoricCharts), _
                    mTradeBuildAPI.StudyLibraryManager, _
                    mTradeBuildAPI.HistoricalDataStoreInput, _
                    mTradeBuildAPI.BarFormatterLibManager, _
                    gMainForm
                    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub StopSelectedTickers()
Const ProcName As String = "StopSelectedTickers"
On Error GoTo Err

Dim lTickers As SelectedTickers
Set lTickers = TickerGrid1.SelectedTickers

TickerGrid1.StopSelectedTickers

Dim lTicker As IMarketDataSource
For Each lTicker In lTickers
    lTicker.Finish
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub updateInstanceSettings()
Const ProcName As String = "updateInstanceSettings"
On Error GoTo Err

If mAppInstanceConfig Is Nothing Then Exit Sub

mAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionMainForm
Select Case Me.WindowState
Case FormWindowStateConstants.vbMaximized
    mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateMaximized
Case FormWindowStateConstants.vbMinimized
    mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateMinimized
Case FormWindowStateConstants.vbNormal
    mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateNormal
    mAppInstanceConfig.SetSetting ConfigSettingMainFormLeft, Me.left / Screen.TwipsPerPixelX
    mAppInstanceConfig.SetSetting ConfigSettingMainFormTop, Me.Top / Screen.TwipsPerPixelY
    mAppInstanceConfig.SetSetting ConfigSettingMainFormWidth, Me.Width / Screen.TwipsPerPixelX
    mAppInstanceConfig.SetSetting ConfigSettingMainFormHeight, Me.Height / Screen.TwipsPerPixelY
End Select

mAppInstanceConfig.SetSetting ConfigSettingFeaturesPanelHidden, CStr(mFeaturesPanelHidden)
mAppInstanceConfig.SetSetting ConfigSettingFeaturesPanelPinned, CStr(mFeaturesPanelPinned)

mAppInstanceConfig.SetSetting ConfigSettingMainFormFeaturesHidden, CStr(mInfoPanelHidden)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


