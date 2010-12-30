VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form fDataCollectorUI 
   Caption         =   "TradeBuild Data Collector Version 2.6"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   3855
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8784
         EndProperty
      EndProperty
   End
   Begin VB.TextBox SecsSinceLastTickText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "The number of seconds since the last tick received"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox TicksPerSecText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "The number of ticks received during the last second"
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton ShowHideMonitorButton 
      Caption         =   "Hide activity monitor"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Shows or hides the activity monitor"
      Top             =   360
      Width           =   1695
   End
   Begin TabDlg.SSTab ActivityMonitor 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      TabHeight       =   494
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Activity"
      TabPicture(0)   =   "fQuoteServerUI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TickerScroll"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TickersContainerPicture"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Log"
      TabPicture(1)   =   "fQuoteServerUI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LogText"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Configuration"
      TabPicture(2)   =   "fQuoteServerUI.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "ConfigNameText"
      Tab(2).Control(2)=   "ConfigDetailsButton"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton ConfigDetailsButton 
         Caption         =   "Details..."
         Height          =   375
         Left            =   -71640
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox ConfigNameText 
         Height          =   285
         Left            =   -74400
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   960
         Width           =   3495
      End
      Begin VB.PictureBox TickersContainerPicture 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2625
         ScaleWidth      =   4785
         TabIndex        =   7
         Top             =   360
         Width           =   4815
         Begin VB.PictureBox TickersPicture 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            ScaleHeight     =   285
            ScaleWidth      =   4815
            TabIndex        =   8
            Top             =   0
            Width           =   4815
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   4
               Left            =   3840
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   0
               Left            =   0
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   1
               Left            =   960
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   2
               Left            =   1920
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   3
               Left            =   2880
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   4
               Left            =   4560
               TabIndex        =   18
               Top             =   0
               Width           =   255
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   0
               Left            =   720
               TabIndex        =   17
               Top             =   0
               Width           =   255
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   3
               Left            =   3600
               TabIndex        =   16
               Top             =   0
               Width           =   255
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   2
               Left            =   2640
               TabIndex        =   15
               Top             =   0
               Width           =   255
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   1
               Left            =   1680
               TabIndex        =   14
               Top             =   0
               Width           =   255
            End
         End
      End
      Begin VB.TextBox LogText 
         Height          =   2655
         Left            =   -74880
         MaxLength       =   65535
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
      Begin VB.VScrollBar TickerScroll 
         Height          =   2700
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Current configuration:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton StartStopButton 
      Caption         =   "Start"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      ToolTipText     =   "Starts or stops data collection"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox ConnectionStatusText 
      BackColor       =   &H8000000F&
      Height          =   255
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Indicates the health of the connection to the realtime data source: green is ok, red is error"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Secs no data"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ticks per sec"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Connection status"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "fDataCollectorUI"
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

Implements BarWriterListener
Implements LogListener
Implements QuoteListener
Implements RawMarketDepthListener
Implements StateChangeListener
Implements TickfileWriterListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "fDataCollectorUI"

Private Const TickerScrollMax As Integer = 32767
Private Const TickerScrollMin As Integer = 0

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type TickerTableEntry
    theTicker               As ticker
    tli                     As TimerListItem
End Type

'================================================================================
' Member variables
'================================================================================

Private WithEvents mDataCollector As DataCollector
Attribute mDataCollector.VB_VarHelpID = -1

Private mTickers() As TickerTableEntry

Private mTimerList As TimerList

Private mLastTickTime As Date

Private WithEvents mClock As Clock
Attribute mClock.VB_VarHelpID = -1

Private mTickCount As Long

Private mActivityMonitorVisible As Boolean

Private mCollectingData As Boolean

Private mAdjustingSize As Boolean
Private mCurrentHeight As Long
Private mCurrentWidth As Long

Private mStartStopButtonInitialLeft As Long

Private mLinesToScroll As Integer

Private mLineSpacing As Integer

Private mStartStopTimePanel As Panel

Private mFormatter As LogFormatter

Private mConfigManager As ConfigManager

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

InitCommonControls
Set mTimerList = GetGlobalTimerList
ReDim mTickers(99) As TickerTableEntry
Set mFormatter = CreateBasicLogFormatter(TimestampTimeOnlyLocal)
GetLogger("log").AddLogListener Me

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

mStartStopButtonInitialLeft = StartStopButton.Left
TickerScroll.Min = TickerScrollMin
TickerScroll.Max = TickerScrollMax
mLineSpacing = ShortNameText(0).Height - Screen.TwipsPerPixelY
mCurrentHeight = Me.Height
mCurrentWidth = Me.Width

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

Select Case UnloadMode
Case QueryUnloadConstants.vbAppTaskManager
Case QueryUnloadConstants.vbAppWindows
Case QueryUnloadConstants.vbFormCode
Case QueryUnloadConstants.vbFormControlMenu
    If mCollectingData Then
        If MsgBox("Please confirm that you wish to stop data collection", _
                    vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Cancel = True
    End If
End Select

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Resize()

Const ProcName As String = "Form_Resize"
On Error GoTo Err

If Me.WindowState = vbMinimized Then Exit Sub
If Me.Height <> mCurrentHeight Then resizeHeight
If Me.Width <> mCurrentWidth Then resizeWidth

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName

End Sub

Private Sub Form_Terminate()
Const ProcName As String = "Form_Terminate"
On Error GoTo Err

TerminateTWUtilities

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

TradeBuild.TradeBuildAPI.ServiceProviders.RemoveAll

LogMessage "Data Collector program exiting"
GetLogger("log").RemoveLogListener Me

TerminateTWUtilities

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' BarWriterListener Interface Members
'================================================================================

Private Sub BarWriterListener_notify(ev As TradeBuild26.WriterEventData)
Dim tk As ticker

Const ProcName As String = "BarWriterListener_notify"
On Error GoTo Err

Select Case ev.Action
Case WriterNotifications.WriterNotReady
    Set tk = ev.Source
    LogMessage "Bar writer not ready for " & _
                tk.Contract.Specifier.localSymbol
Case WriterNotifications.WriterReady
    Set tk = ev.Source
    LogMessage "Bar writer ready for " & _
                tk.Contract.Specifier.localSymbol
Case WriterNotifications.WriterFileCreated
    Set tk = ev.Source
    LogMessage "Writing bars for " & _
                tk.Contract.Specifier.localSymbol & _
                " to " & ev.filename
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' LogListener Interface Members
'================================================================================

Private Sub LogListener_finish()
'nothing to do
Const ProcName As String = "LogListener_finish"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub LogListener_Notify(ByVal logrec As TWUtilities30.LogRecord)

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
LogText.SelText = mFormatter.FormatRecord(logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_ask"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_bid(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_bid"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_high(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_high"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_Low(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_Low"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_openInterest"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_previousClose(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_previousClose"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_sessionOpen(ev As TradeBuild26.QuoteEventData)
Const ProcName As String = "QuoteListener_sessionOpen"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_trade(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_trade"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_volume(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_volume"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' RawMarketDepth Interface Members
'================================================================================

Private Sub RawMarketDepthListener_MarketDepthNotAvailable(ByVal reason As String)

Const ProcName As String = "RawMarketDepthListener_MarketDepthNotAvailable"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub RawMarketDepthListener_resetMarketDepth(ev As TradeBuild26.RawMarketDepthEventData)

Const ProcName As String = "RawMarketDepthListener_resetMarketDepth"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub RawMarketDepthListener_updateMarketDepth(ev As TradeBuild26.RawMarketDepthEventData)
Const ProcName As String = "RawMarketDepthListener_updateMarketDepth"
On Error GoTo Err

processTickEvent ev.Source

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Dim tli As TimerListItem

Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Set tli = ev.Source
If Not tli Is Nothing Then
    If ev.state = TimerListItemStates.TimerListItemStateExpired Then
        switchDataLightOff tli.Data
        Set mTickers(tli.Data).tli = Nothing
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' TickfileWriterListener Interface Members
'================================================================================

Private Sub TickfileWriterListener_notify(ev As TradeBuild26.WriterEventData)
Dim tk As ticker

Const ProcName As String = "TickfileWriterListener_notify"
On Error GoTo Err

Select Case ev.Action
Case WriterNotifications.WriterNotReady
    Set tk = ev.Source
    LogMessage "Tickfile writer not ready for " & _
                    tk.Contract.Specifier.localSymbol
Case WriterNotifications.WriterReady
    Set tk = ev.Source
    LogMessage "Tickfile writer ready for " & _
                    tk.Contract.Specifier.localSymbol
Case WriterNotifications.WriterFileCreated
    Set tk = ev.Source
    LogMessage "Writing tickdata for " & _
                tk.Contract.Specifier.localSymbol & _
                " to " & ev.filename
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub ConfigDetailsButton_Click()
Dim f As New fConfig
Const ProcName As String = "ConfigDetailsButton_Click"
On Error GoTo Err

Set f = New fConfig

f.initialise mConfigManager, True
f.Show vbModeless

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName

End Sub

Private Sub ShowHideMonitorButton_Click()
Const ProcName As String = "ShowHideMonitorButton_Click"
On Error GoTo Err

If mActivityMonitorVisible Then
    hideActivityMonitor
Else
    showActivityMonitor
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub StartStopButton_Click()
Const ProcName As String = "StartStopButton_Click"
On Error GoTo Err

If mCollectingData Then
    stopCollecting "Data collection stopped by user", True
Else
    startCollecting "Data collection started by user"
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub TickerScroll_Change()
Const ProcName As String = "TickerScroll_Change"
On Error GoTo Err

scrollTickers

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mClock Event Handlers
'================================================================================

Private Sub mClock_Tick()
Const ProcName As String = "mClock_Tick"
On Error GoTo Err

TicksPerSecText = mTickCount
TicksPerSecText.Refresh
mTickCount = 0
SecsSinceLastTickText = Format(86400 * (GetTimestamp - mLastTickTime), "0")
SecsSinceLastTickText.Refresh

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mDataCollector Event Handlers
'================================================================================

Private Sub mDataCollector_CollectionStarted()
Dim s As String
Const ProcName As String = "mDataCollector_CollectionStarted"
On Error GoTo Err

If mDataCollector.nextEndTime <> 0 Then
    s = "Collection end: " & _
        FormatDateTime(mDataCollector.nextEndTime, vbShortDate) & " " & _
        FormatDateTime(mDataCollector.nextEndTime, vbShortTime)
End If
If mDataCollector.exitProgramTime <> 0 Then
    s = IIf(s = "", "P", s & "; p") & _
        "rogram exit: " & _
        FormatDateTime(mDataCollector.exitProgramTime, vbShortDate) & " " & _
        FormatDateTime(mDataCollector.exitProgramTime, vbShortTime)
End If

mStartStopTimePanel.Text = s

setStarted

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_CollectionStopped()
Dim s As String
Const ProcName As String = "mDataCollector_CollectionStopped"
On Error GoTo Err

If mDataCollector.nextStartTime <> 0 Then
    s = "Collection start: " & _
        FormatDateTime(mDataCollector.nextStartTime, vbShortDate) & " " & _
        FormatDateTime(mDataCollector.nextStartTime, vbShortTime)
End If
If mDataCollector.exitProgramTime <> 0 Then
    s = IIf(s = "", "P", s & "; p") & _
        "rogram exit: " & _
        FormatDateTime(mDataCollector.exitProgramTime, vbShortDate) & " " & _
        FormatDateTime(mDataCollector.exitProgramTime, vbShortTime)
End If

mStartStopTimePanel.Text = s

setStopped

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_connected()
Const ProcName As String = "mDataCollector_connected"
On Error GoTo Err

ConnectionStatusText.BackColor = vbGreen
StartStopButton.enabled = True

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_connectFailed(ByVal description As String)
Const ProcName As String = "mDataCollector_connectFailed"
On Error GoTo Err

ConnectionStatusText.BackColor = vbRed

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_ConnectionClosed()
Const ProcName As String = "mDataCollector_ConnectionClosed"
On Error GoTo Err

ConnectionStatusText.BackColor = vbRed
StartStopButton.enabled = False

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_Error(ev As TWUtilities30.ErrorEventData)
Const ProcName As String = "mDataCollector_Error"
On Error GoTo Err

LogMessage "Error " & ev.errorCode & ": " & vbCrLf & _
            ev.ErrorMessage, _
            LogLevelSevere

stopCollecting "Closing due to error", False
Unload Me

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_ExitProgram()
Const ProcName As String = "mDataCollector_ExitProgram"
On Error GoTo Err

Unload Me

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_FatalError(ev As TWUtilities30.ErrorEventData)
Const ProcName As String = "mDataCollector_FatalError"
On Error GoTo Err

gHandleFatalError

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_Reconnecting()
Const ProcName As String = "mDataCollector_Reconnecting"
On Error GoTo Err

StartStopButton.enabled = True

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mDataCollector_TickerAdded(ByVal ticker As ticker)
Dim i As Long
Dim index As Long

Const ProcName As String = "mDataCollector_TickerAdded"
On Error GoTo Err

index = ticker.Handle
If index > UBound(mTickers) Then
    ReDim Preserve mTickers(index / 100 * 100 + 99) As TickerTableEntry
End If
Set mTickers(index).theTicker = ticker
Set mTickers(index).tli = Nothing

If index > ShortNameText.UBound Then
    For i = ShortNameText.UBound + 1 To index
        Load ShortNameText(i)
        ShortNameText(i).Left = ShortNameText(i - 5).Left
        ShortNameText(i).Top = ShortNameText(i - 5).Top + mLineSpacing
        ShortNameText(i).ZOrder 0
        If i Mod 5 = 0 Then TickersPicture.Height = ShortNameText(i).Top + ShortNameText(i).Height
        
        Load DataLightLabel(i)
        DataLightLabel(i).Left = DataLightLabel(i - 5).Left
        DataLightLabel(i).Top = DataLightLabel(i - 5).Top + mLineSpacing
        DataLightLabel(i).ZOrder 0
        
        setupTickerScroll
    Next
End If

ShortNameText(index) = ticker.Contract.Specifier.localSymbol
ShortNameText(index).ToolTipText = ticker.Contract.Specifier.localSymbol
ShortNameText(index).Visible = True
DataLightLabel(index).Visible = True

ticker.AddQuoteListener Me
ticker.AddRawMarketDepthListener Me
ticker.AddTickfileWriterListener Me
ticker.AddBarWriterListener Me

Me.Refresh

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal pDataCollector As DataCollector, _
                ByVal pconfigManager As ConfigManager, _
                ByVal configName As String, _
                ByVal noAutoStart As Boolean, _
                ByVal showMonitor As Boolean)

Const ProcName As String = "initialise"
On Error GoTo Err

Set mStartStopTimePanel = StatusBar1.Panels.Item(1)

Set mDataCollector = pDataCollector

Set mConfigManager = pconfigManager
ConfigNameText = configName

If showMonitor Then
    mActivityMonitorVisible = True
Else
    hideActivityMonitor
End If

If noAutoStart Then
    If mDataCollector.exitProgramTime <> 0 Then
        mStartStopTimePanel.Text = "Program exit: " & _
                                    FormatDateTime(mDataCollector.exitProgramTime, vbShortDate) & " " & _
                                    FormatDateTime(mDataCollector.exitProgramTime, vbShortTime)
    End If
Else
Dim s As String
    If mDataCollector.nextStartTime <> 0 Then
        s = "Collection start: " & _
            FormatDateTime(mDataCollector.nextStartTime, vbShortDate) & " " & _
            FormatDateTime(mDataCollector.nextStartTime, vbShortTime)
    End If
    If mDataCollector.exitProgramTime <> 0 Then
        s = IIf(s = "", "P", s & "; p") & _
            "rogram exit: " & _
            FormatDateTime(mDataCollector.exitProgramTime, vbShortDate) & " " & _
            FormatDateTime(mDataCollector.exitProgramTime, vbShortTime)
    End If
    mStartStopTimePanel.Text = s
    If mDataCollector.nextStartTime = 0 Then
        startCollecting "Data collection started automatically"
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearTickers()
Dim i As Long

Const ProcName As String = "clearTickers"
On Error GoTo Err

For i = 0 To ShortNameText.UBound
    ShortNameText(i).Text = ""
    DataLightLabel(i).BackColor = vbButtonFace
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function generateTaskInfo( _
                ByVal en As Enumerator) As String
Dim ts As TaskSummary
Dim s As String

Const ProcName As String = "generateTaskInfo"
On Error GoTo Err

Do While en.MoveNext
    ts = en.Current
    s = s & "Name: " & ts.name & _
        "; Priority: " & ts.priority & _
        "; Start time: " & FormatTimestamp(ts.startTime, TimestampDateAndTimeISO8601) & _
        "; Last run time: " & FormatTimestamp(ts.lastRunTime, TimestampDateAndTimeISO8601) & _
        "; CPU time: " & Format(ts.totalCPUTime, "0.0") & vbCrLf
        
Loop
    
generateTaskInfo = s

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub hideActivityMonitor()
Const ProcName As String = "hideActivityMonitor"
On Error GoTo Err

ShowHideMonitorButton.Caption = "Show activity monitor"
mAdjustingSize = True
Me.Height = Me.Height - ActivityMonitor.Height
mAdjustingSize = False
ActivityMonitor.Visible = False
mActivityMonitorVisible = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
    
Private Sub processTickEvent( _
                pTicker As ticker)
Const ProcName As String = "processTickEvent"
On Error GoTo Err

switchDataLightOn pTicker.Handle
mLastTickTime = GetTimestamp
mTickCount = mTickCount + 1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub resizeHeight()
Dim heightIncrement As Long

Const ProcName As String = "resizeHeight"
On Error GoTo Err

If Not mActivityMonitorVisible And Not mAdjustingSize Then
    Me.Height = mCurrentHeight
    Exit Sub
End If

heightIncrement = Me.Height - mCurrentHeight

If Not mAdjustingSize Then
    If TickersContainerPicture.Height + heightIncrement <= 0 Then
        Me.Height = mCurrentHeight
        Exit Sub
    End If
    
    ActivityMonitor.Height = ActivityMonitor.Height + heightIncrement

    TickersContainerPicture.Height = TickersContainerPicture.Height + heightIncrement

    TickerScroll.Height = TickerScroll.Height + heightIncrement

    LogText.Height = LogText.Height + heightIncrement
    
End If

mCurrentHeight = Me.Height
setupTickerScroll

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub resizeWidth()
Dim widthIncrement As Long

Const ProcName As String = "resizeWidth"
On Error GoTo Err

If Not mActivityMonitorVisible And Not mAdjustingSize Then
    Me.Width = mCurrentWidth
    Exit Sub
End If

If Me.Width <= mStartStopButtonInitialLeft + StartStopButton.Width + 120 Then
    Me.Width = mCurrentWidth
    Exit Sub
End If

widthIncrement = Me.Width - mCurrentWidth

If Not mAdjustingSize Then
    
    ActivityMonitor.Width = ActivityMonitor.Width + widthIncrement

    TickersPicture.Width = TickersPicture.Width + widthIncrement

    TickerScroll.Left = TickerScroll.Left + widthIncrement

    LogText.Width = LogText.Width + widthIncrement

    StartStopButton.Left = StartStopButton.Left + widthIncrement
End If

mCurrentWidth = Me.Width

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub scrollTickers()
Const ProcName As String = "scrollTickers"
On Error GoTo Err

If TickersPicture.Height <= TickersContainerPicture.Height Then
    TickersPicture.Top = 0
ElseIf TickerScroll.Value = TickerScrollMax Then
    TickersPicture.Top = -mLinesToScroll * mLineSpacing
Else
    TickersPicture.Top = -Round((mLinesToScroll / TickerScrollMax) * TickerScroll.Value, 0) * mLineSpacing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setStarted()
Const ProcName As String = "setStarted"
On Error GoTo Err

mCollectingData = True
StartStopButton.Caption = "Stop"
StartStopButton.enabled = True

mLastTickTime = GetTimestamp

Set mClock = GetClock

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setStopped()

Const ProcName As String = "setStopped"
On Error GoTo Err

mCollectingData = False
StartStopButton.Caption = "Start"

ConnectionStatusText.BackColor = vbButtonFace

clearTickers

SecsSinceLastTickText = ""
TicksPerSecText = ""

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub setupTickerScroll()
Dim totalLines As Long
Dim linesPerpage As Single
Dim pagesToScroll As Single

Const ProcName As String = "setupTickerScroll"
On Error GoTo Err

totalLines = (ShortNameText.UBound + 5) / 5
linesPerpage = TickersContainerPicture.Height / mLineSpacing
If totalLines > linesPerpage Then mLinesToScroll = -Int(linesPerpage - totalLines)
pagesToScroll = mLinesToScroll / linesPerpage

If mLinesToScroll > 1 Then
    TickerScroll.SmallChange = (CLng(TickerScrollMax) + mLinesToScroll - 1) / mLinesToScroll
Else
    TickerScroll.SmallChange = TickerScrollMax
End If

If pagesToScroll > 1 Then
    TickerScroll.LargeChange = TickerScrollMax / pagesToScroll
Else
    TickerScroll.LargeChange = TickerScrollMax
End If
TickerScroll.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub showActivityMonitor()
Const ProcName As String = "showActivityMonitor"
On Error GoTo Err

ShowHideMonitorButton.Caption = "Hide activity monitor"
mAdjustingSize = True
Me.Height = Me.Height + ActivityMonitor.Height
mAdjustingSize = False
ActivityMonitor.Visible = True
mActivityMonitorVisible = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub startCollecting( _
                ByVal message As String)
                
Const ProcName As String = "startCollecting"
On Error GoTo Err

LogMessage message

mDataCollector.startCollection

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub stopCollecting( _
                ByVal message As String, _
                ByVal confirm As Boolean)
Const ProcName As String = "stopCollecting"
On Error GoTo Err

If confirm Then
    If MsgBox("Please confirm that you wish to stop data collection", _
                vbYesNo + vbDefaultButton2 + vbQuestion) <> vbYes Then Exit Sub
End If

LogMessage message

mDataCollector.stopCollection

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub switchDataLightOn( _
                ByVal index As Long)
Const ProcName As String = "switchDataLightOn"
On Error GoTo Err

If Not mActivityMonitorVisible Then Exit Sub

If Not mTickers(index).tli Is Nothing Then
    mTimerList.Remove mTickers(index).tli
    mTickers(index).tli.RemoveStateChangeListener Me
End If

Set mTickers(index).tli = mTimerList.Add(index, 200, ExpiryTimeUnitMilliseconds)
mTickers(index).tli.AddStateChangeListener Me

DataLightLabel(index).BackColor = vbGreen
DataLightLabel(index).Refresh
ConnectionStatusText.BackColor = vbGreen

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub switchDataLightOff( _
                ByVal index As Long)
Const ProcName As String = "switchDataLightOff"
On Error GoTo Err

DataLightLabel(index).BackColor = vbButtonFace
DataLightLabel(index).Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

