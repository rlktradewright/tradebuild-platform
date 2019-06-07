VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#340.0#0"; "TradingUI27.ocx"
Begin VB.Form PositionForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox LogText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3240
      Width           =   10095
   End
   Begin TradingUI27.ExecutionsSummary ExecutionsSummary1 
      Height          =   3015
      Left            =   3120
      TabIndex        =   17
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
   End
   Begin VB.TextBox PendingSizeText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton ClosePositionButton 
      Caption         =   "Close Position"
      Height          =   495
      Left            =   10320
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TradeDrawdownText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox TradeMaxProfitText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox TradeProfitText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox SessionDrawdownText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox SessionMaxProfitText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox SessionProfitText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox PositionSizeText 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Pending size"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Trade drawdown"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Trade max profit"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Trade profit"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Session drawdown"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Session max profit"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Session profit"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Position size"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PositionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IProfitListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "PositionForm"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mPositionManager                 As PositionManager
Attribute mPositionManager.VB_VarHelpID = -1

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ClosePositionButton_Click()
Const ProcName As String = "ClosePositionButton_Click"
On Error GoTo Err

mPositionManager.ClosePositions

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IProfitListener Interface Members
'@================================================================================

Private Sub IProfitListener_NotifyProfit(ev As ProfitEventData)
Dim lProfitTypes As ProfitTypes
lProfitTypes = ev.ProfitTypes

Dim lPositionManager As PositionManager
Set lPositionManager = ev.Source

If lProfitTypes And ProfitTypeSessionProfit Then _
    SessionProfitText.Text = lPositionManager.Profit
If lProfitTypes And ProfitTypeSessionMaxProfit Then _
    SessionMaxProfitText.Text = lPositionManager.MaxProfit
If lProfitTypes And ProfitTypeSessionDrawdown Then _
    SessionDrawdownText.Text = lPositionManager.Drawdown
If lProfitTypes And ProfitTypeTradeProfit Then _
    TradeProfitText.Text = lPositionManager.ProfitThisTrade
If lProfitTypes And ProfitTypeTradeMaxProfit Then _
    TradeMaxProfitText.Text = lPositionManager.MaxProfitThisTrade
If lProfitTypes And ProfitTypeTradeDrawdown Then _
    TradeDrawdownText.Text = lPositionManager.DrawdownThisTrade
End Sub

'@================================================================================
' mPositionManager Event Handlers
'@================================================================================

Private Sub mPositionManager_Change(ev As ChangeEventData)
Dim lChangeType As PositionManagerChangeTypes
Const ProcName As String = "mPositionManager_Change"
On Error GoTo Err

lChangeType = ev.ChangeType

Select Case lChangeType
Case PositionSizeChanged
    PositionSizeText.Text = mPositionManager.PositionSize
    PendingSizeText.Text = mPositionManager.PendingPositionSize
Case ProviderReadinessChanged

Case PositionClosed

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mPositionManager_OrderError(ByVal pOrderId As String, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "mPositionManager_OrderError"
On Error GoTo Err

Dim s As String
s = FormatTimestamp(Now, TimestampDateAndTimeISO8601 + TimestampNoMillisecs) & "  " & _
    "Error " & CStr(pErrorCode) & " (order id=" & pOrderId & "): " & pErrorMsg

LogMessage s, LogLevelWarning
LogText.Text = LogText.Text & vbCrLf & s

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pPositionManager As PositionManager, _
                ByVal pTheme As ITheme)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mPositionManager = pPositionManager
ExecutionsSummary1.MonitorExecutions mPositionManager.Executions

Dim lContract As IContract
Set lContract = mPositionManager.ContractFuture.Value
Me.Caption = "Position for " & lContract.Specifier.LocalSymbol
PositionSizeText.Text = mPositionManager.PositionSize
PendingSizeText.Text = mPositionManager.PendingPositionSize
mPositionManager.AddProfitListener Me

Set mTheme = pTheme
Me.BackColor = mTheme.BaseColor
gApplyTheme mTheme, Me.Controls

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




